/**
 * BACKEND APPS SCRIPT — Inventaris Alat & Aksesoris, Instalasi Anestesi RSUA
 *
 * Kontrak dengan index.html:
 *   GET  ?callback=fn                 -> fn({ok:true, alat:[...], aksesoris:[...]})   (JSONP)
 *   POST payload={"action":"save","alat":{...},"aksesoris":[...]}
 *   POST payload={"action":"delete","id":"..."}
 *
 * Data disimpan di dua tab yang dibuat otomatis: "Alat" dan "Aksesoris".
 * Disarankan dipasang di spreadsheet TERPISAH (mis. "Database Inventaris Anestesi"),
 * bukan di file laporan 5 tab, supaya file laporan tetap bersih.
 */

// Kosongkan jika script ini dibuat lewat Ekstensi > Apps Script di dalam spreadsheet-nya.
// Isi dengan ID spreadsheet bila script berdiri sendiri (script.google.com).
const SPREADSHEET_ID = '';

const ALAT_SHEET = 'Alat';
const AKS_SHEET = 'Aksesoris';
const ALAT_COLS = ['id', 'nama', 'merek', 'model', 'no_seri', 'no_inventaris', 'lokasi', 'kondisi',
  'kalibrasi_terakhir', 'kalibrasi_berikutnya', 'catatan',
  'created_by', 'created_at', 'updated_by', 'updated_at'];
const AKS_COLS = ['id', 'alat_id', 'nama', 'jumlah', 'status', 'catatan'];
const MAX_ALAT_AKS = 200;   // batas aksesoris per alat

/* ============================ ENDPOINT ============================ */

function doGet(e) {
  var out;
  try {
    out = {
      ok: true,
      alat: readTable_(getSheet_(ALAT_SHEET, ALAT_COLS), ALAT_COLS),
      aksesoris: readTable_(getSheet_(AKS_SHEET, AKS_COLS), AKS_COLS)
    };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  var json = JSON.stringify(out).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z_$][\w$.]*$/.test(cb)) {     // nama callback divalidasi (cegah injeksi)
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(25000);            // cegah dua penyimpanan bersamaan saling menimpa
    locked = true;
    var raw = (e && e.parameter && e.parameter.payload) ||
              (e && e.postData && e.postData.contents) || '{}';
    var p = JSON.parse(raw);
    if (p.action === 'save') saveAlat_(p);
    else if (p.action === 'delete') deleteAlat_(p.id);
    else throw new Error('Aksi tidak dikenal.');
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (_) {} }
  }
}

/* ============================ OPERASI ============================ */

function saveAlat_(p) {
  var a = p.alat || {};
  var id = clean_(a.id);
  if (!id) throw new Error('ID alat kosong.');
  if (!clean_(a.nama)) throw new Error('Nama alat wajib diisi.');

  var alatSheet = getSheet_(ALAT_SHEET, ALAT_COLS);
  var aksSheet = getSheet_(AKS_SHEET, AKS_COLS);

  var rec = {};
  ALAT_COLS.forEach(function (c) { rec[c] = clean_(a[c]); });
  rec.id = id;

  var alat = readTable_(alatSheet, ALAT_COLS);
  var i = alat.findIndex(function (x) { return x.id === id; });
  if (i >= 0) alat[i] = rec; else alat.push(rec);
  writeTable_(alatSheet, ALAT_COLS, alat);

  // Ganti seluruh aksesoris milik alat ini dengan kiriman terbaru.
  var aks = readTable_(aksSheet, AKS_COLS).filter(function (k) { return k.alat_id !== id; });
  (p.aksesoris || []).slice(0, MAX_ALAT_AKS).forEach(function (k) {
    var nama = clean_(k.nama);
    if (!nama) return;
    aks.push({
      id: clean_(k.id) || Utilities.getUuid(),
      alat_id: id,
      nama: nama,
      jumlah: Math.max(0, Number(k.jumlah) || 0),
      status: clean_(k.status) || 'Lengkap',
      catatan: clean_(k.catatan)
    });
  });
  writeTable_(aksSheet, AKS_COLS, aks);
}

function deleteAlat_(id) {
  id = clean_(id);
  if (!id) throw new Error('ID alat kosong.');
  var alatSheet = getSheet_(ALAT_SHEET, ALAT_COLS);
  var aksSheet = getSheet_(AKS_SHEET, AKS_COLS);
  writeTable_(alatSheet, ALAT_COLS,
    readTable_(alatSheet, ALAT_COLS).filter(function (x) { return x.id !== id; }));
  writeTable_(aksSheet, AKS_COLS,
    readTable_(aksSheet, AKS_COLS).filter(function (k) { return k.alat_id !== id; }));
}

/* ============================ UTIL SHEET ============================ */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v) {
  return String(v == null ? '' : v)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .trim()
    .slice(0, 1000);
}

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// Ambil tab; jika belum ada, buat lengkap dengan header. Kolom teks diformat "teks biasa"
// agar nomor seri / tanggal tidak diubah otomatis oleh Google Sheets.
function getSheet_(name, cols) {
  var ss = getSpreadsheet_();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan. Isi SPREADSHEET_ID di Code.gs.');
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.setFrozenRows(1);
    var rows = sh.getMaxRows();
    cols.forEach(function (c, i) {
      sh.getRange(2, i + 1, Math.max(1, rows - 1), 1).setNumberFormat(c === 'jumlah' ? '0' : '@');
    });
  }
  return sh;
}

function readTable_(sh, cols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return vals
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      var o = {};
      cols.forEach(function (c, i) {
        var v = r[i];
        if (c === 'jumlah') o[c] = Number(v) || 0;
        else if (v instanceof Date) o[c] = v.toISOString();
        else o[c] = String(v == null ? '' : v);
      });
      return o;
    });
}

function writeTable_(sh, cols, objs) {
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (!objs.length) return;
  var need = objs.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(2, 1, objs.length, cols.length).setValues(
    objs.map(function (o) { return cols.map(function (c) { return o[c] == null ? '' : o[c]; }); })
  );
}
