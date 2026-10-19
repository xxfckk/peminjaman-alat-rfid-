// ===== Peminjaman Alat RPL/IoT (v3) - Code.gs =====
// Tab: "Peminjam" (ID Tag|Nama|Kelas)
//      "Barang"   (ID Card|Nama Barang|Kondisi|Status)
//      "Log"      (Waktu Pinjam|Nama|Kelas|ID Tag|ID Card|Nama Barang|Kondisi|Waktu Kembali|Status)
// File lain di proyek ini: index.html (nama harus "index")

// Kosongkan jika script dibuka lewat Extensions > Apps Script di dalam Spreadsheet-nya.
// Jika script berdiri sendiri, isi dengan ID spreadsheet (teks panjang di URL antara /d/ dan /edit).
var SS_ID = '1IGP4biqdp-k5OrBDppLq_Pc6vBioJSIJ83pTrWX3hPQ'; // spreadsheet lama Anda
var TZ = 'Asia/Jakarta';

function ss_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// JALANKAN SEKALI: membuat 3 tab + header + format teks otomatis
function siapkan() {
  var def = {
    'Peminjam': ['ID Tag', 'Nama', 'Kelas'],
    'Barang': ['ID Card', 'Nama Barang', 'Kondisi', 'Status'],
    'Log': ['Waktu Pinjam', 'Nama', 'Kelas', 'ID Tag', 'ID Card', 'Nama Barang', 'Kondisi', 'Waktu Kembali', 'Status']
  };
  var ss = ss_();
  for (var n in def) {
    var s = ss.getSheetByName(n) || ss.insertSheet(n);
    var w = def[n].length;
    s.getRange(1, 1, s.getMaxRows(), w).setNumberFormat('@');
    s.getRange(1, 1, 1, w).setValues([def[n]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
}

// ---------- Pintu masuk: web (tanpa action) & ESP32 (dengan action) ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Pinjam Alat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
  var out;
  try {
    if (p.action === 'ping') { beat(); out = { ok: true }; }
    else if (p.action === 'scan') out = scan(p.uid);
    else if (p.action === 'transaksi') out = transaksi(p.tag, p.card);
    else out = { ok: false, msg: 'aksi tidak dikenal' };
  } catch (err) {
    out = { ok: false, msg: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Helper ----------
function sh(nama) {
  var s = ss_().getSheetByName(nama);
  if (!s) throw new Error('Tab "' + nama + '" tidak ditemukan');
  return s;
}
function norm(v) { return String(v).trim().toUpperCase(); }
function str(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm:ss') : String(v);
}
function cari(sheet, uid) {
  if (!uid) return null;
  var d = sheet.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (norm(d[i][0]) === norm(uid)) return { row: i + 1, data: d[i] };
  }
  return null;
}
function setProp(k, v) { PropertiesService.getScriptProperties().setProperty(k, JSON.stringify(v)); }
function getProp(k) {
  var s = PropertiesService.getScriptProperties().getProperty(k);
  return s ? JSON.parse(s) : null;
}
function beat() { setProp('ping', Date.now()); }          // tanda ESP32 hidup
function bump() { setProp('rev', (getProp('rev') || 0) + 1); } // tanda data berubah

// ---------- Dipanggil ESP32 ----------
function scan(uid) {
  beat();
  if (!uid) return { ok: false, type: 'unknown' };
  var r;
  var a = cari(sh('Peminjam'), uid);
  if (a) r = { ok: true, type: 'tag', nama: a.data[1], kelas: a.data[2] };
  else {
    var b = cari(sh('Barang'), uid);
    if (b) r = { ok: true, type: 'card', nama: b.data[1], status: b.data[3] };
    else r = { ok: false, type: 'unknown' };
  }
  setProp('scan', { uid: norm(uid), type: r.type, nama: r.nama || '', t: Date.now() });
  return r;
}

function transaksi(tag, card) {
  beat();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var a = cari(sh('Peminjam'), tag);
    var b = cari(sh('Barang'), card);
    if (!a || !b) return { ok: false, msg: 'ID tidak dikenal' };

    var wkt = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    var log = sh('Log');
    var barang = sh('Barang');

    if (String(b.data[3]).trim() === 'Dipinjam') {
      var d = log.getDataRange().getValues();
      for (var i = d.length - 1; i >= 1; i--) {
        if (norm(d[i][4]) === norm(card) && String(d[i][8]).trim() === 'Dipinjam') {
          log.getRange(i + 1, 8, 1, 2).setNumberFormat('@');
          log.getRange(i + 1, 8, 1, 2).setValues([[wkt, 'Selesai']]);
          break;
        }
      }
      barang.getRange(b.row, 4).setValue('Tersedia');
      bump();
      return { ok: true, mode: 'kembali', nama: a.data[1], barang: b.data[1] };
    }

    ensureLog(log);
    var due = Utilities.formatDate(new Date(Date.now() + getDurasi() * 60000), TZ, 'dd/MM/yyyy HH:mm:ss');
    var r = log.getLastRow() + 1;
    if (r > log.getMaxRows()) log.insertRowsAfter(log.getMaxRows(), r - log.getMaxRows());
    log.getRange(r, 1, 1, 10).setNumberFormat('@');
    log.getRange(r, 1, 1, 10).setValues([[
      wkt, a.data[1], a.data[2], norm(tag), norm(card),
      b.data[1], b.data[2], '', 'Dipinjam', due
    ]]);
    barang.getRange(b.row, 4).setValue('Dipinjam');
    bump();
    return { ok: true, mode: 'pinjam', nama: a.data[1], barang: b.data[1] };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Dipanggil website (google.script.run) ----------
function apiStatus() {
  var pg = getProp('ping'), sc = getProp('scan'), now = Date.now();
  return {
    rev: getProp('rev') || 0,
    now: now,
    espAge: pg ? Math.round((now - pg) / 1000) : null,
    scan: sc ? { uid: sc.uid, type: sc.type, nama: sc.nama, age: Math.round((now - sc.t) / 1000) } : null
  };
}

function apiData() {
  var rev = getProp('rev') || 0;
  var pem = [], bar = [], log = [], aktif = [];
  sh('Peminjam').getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]).trim() !== '') pem.push({ id: norm(r[0]), nama: str(r[1]), kelas: str(r[2]) });
  });
  sh('Barang').getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]).trim() !== '') bar.push({ id: norm(r[0]), nama: str(r[1]), kondisi: str(r[2]), status: String(r[3]).trim() || 'Tersedia' });
  });
  var ld = sh('Log').getDataRange().getValues();
  for (var i = 1; i < ld.length; i++) {
    var r = ld[i];
    if (String(r[0]).trim() === '') continue;
    var o = { row: i + 1, waktu: str(r[0]), nama: str(r[1]), kelas: str(r[2]), card: norm(r[4]), barang: str(r[5]),
              kembali: str(r[7]), batas: r.length > 9 ? str(r[9]) : '', status: String(r[8]).trim() };
    log.push(o);
    if (o.status === 'Dipinjam') aktif.push(o);
  }
  log.reverse(); aktif.reverse();
  return { rev: rev, now: Date.now(), durasi: getDurasi(), pem: pem, bar: bar, log: log.slice(0, 80), aktif: aktif };
}

function mutate(fn) {
  var l = LockService.getScriptLock();
  l.waitLock(10000);
  try { var r = fn(); bump(); return r; } finally { l.releaseLock(); }
}
function dupe(id) { return cari(sh('Peminjam'), id) || cari(sh('Barang'), id); }

function apiTambahPeminjam(id, nama, kelas) {
  id = norm(id); nama = String(nama).trim(); kelas = String(kelas).trim();
  if (!id || !nama) throw new Error('ID dan nama wajib diisi');
  return mutate(function () {
    if (dupe(id)) throw new Error('ID sudah terdaftar');
    var s = sh('Peminjam'), r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[id, nama, kelas]]);
    return { ok: true };
  });
}

function apiTambahBarang(id, nama, kondisi) {
  id = norm(id); nama = String(nama).trim(); kondisi = String(kondisi).trim() || 'Baik';
  if (!id || !nama) throw new Error('ID dan nama wajib diisi');
  return mutate(function () {
    if (dupe(id)) throw new Error('ID sudah terdaftar');
    var s = sh('Barang'), r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 4).setNumberFormat('@').setValues([[id, nama, kondisi, 'Tersedia']]);
    return { ok: true };
  });
}

function apiHapus(jenis, id) {
  return mutate(function () {
    var s = sh(jenis === 'barang' ? 'Barang' : 'Peminjam');
    var f = cari(s, id);
    if (!f) throw new Error('Data tidak ditemukan');
    if (jenis === 'barang' && String(f.data[3]).trim() === 'Dipinjam') throw new Error('Barang masih dipinjam');
    try { s.deleteRow(f.row); }
    catch (e) { s.getRange(f.row, 1, 1, s.getLastColumn()).clearContent(); } // jika sisa baris tinggal header
    return { ok: true };
  });
}

// Untuk tes: pilih "tes" lalu Jalankan, lihat Log eksekusi
function tes() { Logger.log(JSON.stringify(apiData())); }

// ---------- Fitur edit & batas waktu ----------
function getDurasi() { return getProp('durasi') || 60; } // menit
function ensureLog(log) {
  if (log.getMaxColumns() < 10) log.insertColumnsAfter(log.getMaxColumns(), 10 - log.getMaxColumns());
  if (String(log.getRange(1, 10).getValue()).trim() === '') log.getRange(1, 10).setValue('Batas Waktu');
}
function apiSetDurasi(menit) {
  menit = Math.round(Number(menit));
  if (!(menit >= 1 && menit <= 10080)) throw new Error('Isi 1 - 10080 menit');
  return mutate(function () { setProp('durasi', menit); return { ok: true }; });
}
function apiUbahBarang(id, nama, kondisi) {
  nama = String(nama).trim(); kondisi = String(kondisi).trim() || 'Baik';
  if (!nama) throw new Error('Nama wajib diisi');
  return mutate(function () {
    var s = sh('Barang'), f = cari(s, id);
    if (!f) throw new Error('Barang tidak ditemukan');
    s.getRange(f.row, 2, 1, 2).setNumberFormat('@').setValues([[nama, kondisi]]);
    return { ok: true };
  });
}
function apiUbahPeminjam(id, nama, kelas) {
  nama = String(nama).trim(); kelas = String(kelas).trim();
  if (!nama) throw new Error('Nama wajib diisi');
  return mutate(function () {
    var s = sh('Peminjam'), f = cari(s, id);
    if (!f) throw new Error('Peminjam tidak ditemukan');
    s.getRange(f.row, 2, 1, 2).setNumberFormat('@').setValues([[nama, kelas]]);
    return { ok: true };
  });
}
function apiUbahBatas(row, card, menit) {
  menit = Math.round(Number(menit));
  if (!(menit >= 1 && menit <= 10080)) throw new Error('Isi 1 - 10080 menit');
  return mutate(function () {
    var log = sh('Log'); ensureLog(log);
    var v = log.getRange(Number(row), 1, 1, 10).getValues()[0];
    if (norm(v[4]) !== norm(card) || String(v[8]).trim() !== 'Dipinjam') throw new Error('Data sudah berubah, muat ulang');
    var due = Utilities.formatDate(new Date(Date.now() + menit * 60000), TZ, 'dd/MM/yyyy HH:mm:ss');
    log.getRange(Number(row), 10).setNumberFormat('@').setValue(due);
    return { ok: true };
  });
}
