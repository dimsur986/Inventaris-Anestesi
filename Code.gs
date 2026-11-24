// ===== Backend Inventaris Alat Anestesi — data dari web tersusun rapi di Google Sheet =====
// Akses publik tanpa login sesuai permintaan. Siapa pun yang mengetahui URL Web App dapat membaca, menambah, mengubah, dan menghapus data.
const SS_ID = '1qSm1t-602jZX2Btez3ys9fsDGTaih2D-_nYTxwn_Pwc';
const TZ    = 'Asia/Jakarta';
const NAVY = '#002060', YEL = '#f5b800', LINE = '#d9e0ec', SOFT = '#f5f7fb';
const C = { ok: ['#e4f5ea', '#17803d'], warn: ['#fff0c2', '#8a5d00'], bad: ['#fde8e6', '#b42318'] };
const KONDISI = ['Baik', 'Perlu servis', 'Dalam perbaikan', 'Rusak', 'Tidak aktif'];
const STATUS_AKS = ['Lengkap', 'Kurang', 'Rusak', 'Hilang'];

// Urutan kolom = urutan di sheet (jangan diubah). Label = judul kolom yang tampil.
const H_ALAT = ['id','nama','merek','model','no_seri','no_inventaris','lokasi','kondisi','kalibrasi_terakhir','kalibrasi_berikutnya','catatan','updated_at'];
const L_ALAT = ['ID','Nama Alat','Merek','Model','No. Seri','No. Inventaris','Lokasi','Kondisi','Kalibrasi Terakhir','Kalibrasi Berikutnya','Catatan','Terakhir Diubah'];
const H_AKS  = ['id','alat_id','nama','jumlah','status','catatan','nama_alat'];
const L_AKS  = ['ID','ID Alat','Nama Aksesoris','Jumlah','Status','Catatan','Milik Alat'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SS_ID); }

function sh_(name, labels) {
  const ss = ss_();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(labels);
  return sh;
}

// Dibaca berdasarkan posisi kolom, jadi judul kolom boleh berupa kalimat biasa
function readAll_(sh, keys) {
  const v = sh.getDataRange().getValues();
  v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    keys.forEach((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
      o[k] = x === undefined ? '' : x;
    });
    return o;
  });
}
const alat_ = () => readAll_(sh_('Alat', L_ALAT), H_ALAT);
const aks_  = () => readAll_(sh_('Aksesoris', L_AKS), H_AKS);

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function parseD_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? Utilities.parseDate(s, TZ, 'yyyy-MM-dd') : '';
}

function doGet(e) {
  try {
    return out_({ ok: true, alat: alat_(), aksesoris: aks_() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sA = sh_('Alat', L_ALAT), sK = sh_('Aksesoris', L_AKS);

    if (b.action === 'save') {
      const a = b.alat;
      a.updated_at = new Date();
      a.kalibrasi_terakhir = parseD_(a.kalibrasi_terakhir);
      a.kalibrasi_berikutnya = parseD_(a.kalibrasi_berikutnya);
      const row = H_ALAT.map(k => a[k] === undefined ? '' : a[k]);
      const ids = sA.getRange(1, 1, sA.getLastRow(), 1).getValues().map(r => String(r[0]));
      const idx = ids.indexOf(String(a.id));
      if (idx > 0) sA.getRange(idx + 1, 1, 1, H_ALAT.length).setValues([row]);
      else sA.appendRow(row);
      delAks_(sK, a.id);
      (b.aksesoris || []).forEach(k => { k.nama_alat = a.nama; sK.appendRow(H_AKS.map(h => k[h] === undefined ? '' : k[h])); });
      safeRapi_();
      return out_({ ok: true });
    }
    if (b.action === 'delete') {
      const ids = sA.getRange(1, 1, sA.getLastRow(), 1).getValues().map(r => String(r[0]));
      const idx = ids.indexOf(String(b.id));
      if (idx > 0) sA.deleteRow(idx + 1);
      delAks_(sK, b.id);
      safeRapi_();
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function delAks_(sK, alatId) {
  const v = sK.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][1]) === String(alatId)) sK.deleteRow(i + 1);
}

// ================= PENATAAN SHEET =================
function safeRapi_() {
  try { rapikanData_(); rebuildRekap_(); } catch (err) { Logger.log('Penataan gagal: ' + err); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Inventaris').addItem('Rapikan & Perbarui Rekap', 'rebuildRekap').addToUi();
}

// Jalankan manual dari editor / menu: merapikan semua sheet
function rebuildRekap() { rapikanData_(); rebuildRekap_(); }

function styleDb_(sh, o) {
  const nc = o.labels.length, nr = Math.max(sh.getMaxRows(), 2), last = sh.getLastRow();
  sh.getRange(1, 1, 1, nc).setValues([o.labels]).setFontColor('#ffffff').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 34); sh.setFrozenRows(1); sh.setHiddenGridlines(true);

  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, nr, nc).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(NAVY).setFirstRowColor('#ffffff').setSecondRowColor(SOFT);

  const body = sh.getRange(2, 1, nr - 1, nc);
  body.setFontSize(10).setFontColor(NAVY).setVerticalAlignment('top').setWrap(true);
  o.text.forEach(c => sh.getRange(2, c, nr - 1, 1).setNumberFormat('@'));
  o.date.forEach(c => sh.getRange(2, c, nr - 1, 1).setNumberFormat('dd/MM/yyyy').setHorizontalAlignment('center'));
  o.center.forEach(c => sh.getRange(2, c, nr - 1, 1).setHorizontalAlignment('center'));
  if (last > 1) sh.getRange(1, 1, last, nc).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);

  o.widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.showColumns(1, nc);
  o.hide.forEach(c => sh.hideColumns(c));

  const f = sh.getFilter(); if (f) f.remove();
  if (last > 1) sh.getRange(1, 1, last, nc).createFilter();

  // dropdown + warna otomatis untuk kolom status
  const col = sh.getRange(2, o.statusCol, nr - 1, 1);
  col.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(o.list, true).setAllowInvalid(true).build())
     .setHorizontalAlignment('center').setFontWeight('bold');
  const rule = (txt, k) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt)
    .setBackground(C[k][0]).setFontColor(C[k][1]).setRanges([col]).build();
  sh.setConditionalFormatRules(o.rules.map(r => rule(r[0], r[1])));
}

function rapikanData_() {
  const sA = sh_('Alat', L_ALAT), sK = sh_('Aksesoris', L_AKS);

  // isi nama alat pada aksesoris lama yang belum punya
  if (sK.getLastRow() > 1) {
    const nm = {}; alat_().forEach(a => nm[String(a.id)] = a.nama);
    const rg = sK.getRange(2, 1, sK.getLastRow() - 1, H_AKS.length), v = rg.getValues();
    v.forEach(r => { if (!r[6] && nm[String(r[1])]) r[6] = nm[String(r[1])]; });
    rg.setValues(v);
  }

  styleDb_(sA, {
    labels: L_ALAT, widths: [90, 200, 110, 110, 120, 120, 120, 115, 110, 120, 230, 125],
    text: [1, 5, 6], date: [9, 10], center: [8], hide: [1], statusCol: 8, list: KONDISI,
    rules: [['Baik', 'ok'], ['Perlu servis', 'warn'], ['Dalam perbaikan', 'warn'], ['Rusak', 'bad']]
  });
  sA.getRange(2, 12, Math.max(sA.getMaxRows() - 1, 1), 1).setNumberFormat('dd/MM/yyyy HH:mm').setHorizontalAlignment('center');

  styleDb_(sK, {
    labels: L_AKS, widths: [90, 90, 220, 70, 100, 240, 200],
    text: [1, 2], date: [], center: [4, 5], hide: [1, 2], statusCol: 5, list: STATUS_AKS,
    rules: [['Lengkap', 'ok'], ['Kurang', 'warn'], ['Rusak', 'bad'], ['Hilang', 'bad']]
  });
}

// ================= REKAP =================
function toDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ''));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}

function kalStatus_(next, today) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(next || ''));
  if (!m) return { t: '-', k: '' };
  const diff = Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - today) / 864e5);
  if (diff < 0)   return { t: 'Lewat ' + (-diff) + ' hari', k: 'bad', lewat: true };
  if (diff <= 30) return { t: diff + ' hari lagi', k: 'warn', segera: true };
  return { t: 'Aman', k: 'ok' };
}

function rebuildRekap_() {
  const kondK = k => k === 'Baik' ? 'ok' : (k === 'Perlu servis' || k === 'Dalam perbaikan') ? 'warn' : k === 'Rusak' ? 'bad' : '';
  const ss = ss_();
  const alat = alat_().sort((x, y) => String(x.nama).localeCompare(String(y.nama)));
  const aks = aks_();
  const aksOf = id => aks.filter(k => String(k.alat_id) === String(id));

  const now = new Date();
  const t = Utilities.formatDate(now, TZ, 'yyyy-MM-dd').split('-');
  const today = Date.UTC(+t[0], +t[1] - 1, +t[2]);

  let sh = ss.getSheetByName('Rekap');
  if (!sh) sh = ss.insertSheet('Rekap', 0);
  const old = sh.getFilter(); if (old) old.remove();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);

  const HEAD = ['No', 'Nama Alat', 'Merek', 'Model', 'No. Inventaris', 'No. Seri', 'Lokasi', 'Kondisi',
    'Kalibrasi Terakhir', 'Kalibrasi Berikutnya', 'Status Kalibrasi', 'Jml Aksesoris', 'Aks. Bermasalah',
    'Daftar Aksesoris', 'Catatan'];
  const NC = HEAD.length, HR = 7, n = alat.length;

  sh.getRange(1, 1, 1, NC).merge().setValue('REKAP INVENTARIS ALAT — INSTALASI ANESTESI RSUA')
    .setBackground(NAVY).setFontColor('#ffffff').setFontSize(15).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);
  sh.getRange(2, 1, 1, NC).merge()
    .setValue('Diperbarui otomatis: ' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm') + ' WIB')
    .setBackground(YEL).setFontColor(NAVY).setFontWeight('bold').setFontSize(10);

  const kc = k => alat.filter(a => a.kondisi === k).length;
  const stats = [
    ['Total alat', n], ['Kondisi baik', kc('Baik')], ['Perlu servis', kc('Perlu servis')],
    ['Dalam perbaikan', kc('Dalam perbaikan')], ['Rusak', kc('Rusak')], ['Tidak aktif', kc('Tidak aktif')],
    ['Kalibrasi lewat', alat.filter(a => kalStatus_(a.kalibrasi_berikutnya, today).lewat).length],
    ['Kalibrasi ≤ 30 hari', alat.filter(a => kalStatus_(a.kalibrasi_berikutnya, today).segera).length],
    ['Aksesoris bermasalah', aks.filter(k => k.status !== 'Lengkap').length]
  ];
  sh.getRange(4, 2, 1, stats.length).setValues([stats.map(s => s[0])]).setFontSize(9).setFontColor('#55648a')
    .setFontWeight('bold').setBackground(SOFT).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(5, 2, 1, stats.length).setValues([stats.map(s => s[1])]).setFontSize(20).setFontWeight('bold')
    .setFontColor(NAVY).setBackground('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(4, 2, 2, stats.length).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 34); sh.setRowHeight(5, 42);
  if (stats[4][1] > 0) sh.getRange(5, 6).setFontColor(C.bad[1]);
  if (stats[6][1] > 0) sh.getRange(5, 8).setFontColor(C.bad[1]);
  if (stats[7][1] > 0) sh.getRange(5, 9).setFontColor(C.warn[1]);
  if (stats[8][1] > 0) sh.getRange(5, 10).setFontColor(C.warn[1]);

  sh.getRange(HR, 1, 1, NC).setValues([HEAD]).setBackground(NAVY).setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(HR, 36); sh.setFrozenRows(HR);

  if (n > 0) {
    const vals = [], bgs = [], fcs = [];
    alat.forEach((a, i) => {
      const ak = aksOf(a.id), ks = kalStatus_(a.kalibrasi_berikutnya, today);
      vals.push([i + 1, a.nama, a.merek, a.model, a.no_inventaris, a.no_seri, a.lokasi, a.kondisi,
        toDate_(a.kalibrasi_terakhir), toDate_(a.kalibrasi_berikutnya), ks.t, ak.length,
        ak.filter(k => k.status !== 'Lengkap').length,
        ak.map(k => '• ' + k.nama + ' (' + k.jumlah + ', ' + k.status + ')').join('\n'), a.catatan]);
      const base = i % 2 ? SOFT : '#ffffff';
      const bg = new Array(NC).fill(base), fc = new Array(NC).fill(NAVY);
      const kk = kondK(a.kondisi);
      if (kk) { bg[7] = C[kk][0]; fc[7] = C[kk][1]; }
      if (ks.k) { bg[10] = C[ks.k][0]; fc[10] = C[ks.k][1]; }
      if (ak.some(k => k.status !== 'Lengkap')) { bg[12] = C.warn[0]; fc[12] = C.warn[1]; }
      bgs.push(bg); fcs.push(fc);
    });
    sh.getRange(HR + 1, 5, n, 2).setNumberFormat('@');
    sh.getRange(HR + 1, 1, n, NC).setValues(vals).setBackgrounds(bgs).setFontColors(fcs)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true)
      .setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(HR + 1, 1, n, 1).setHorizontalAlignment('center');
    sh.getRange(HR + 1, 2, n, 1).setFontWeight('bold');
    sh.getRange(HR + 1, 8, n, 1).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(HR + 1, 9, n, 2).setNumberFormat('dd/MM/yyyy').setHorizontalAlignment('center');
    sh.getRange(HR + 1, 11, n, 3).setHorizontalAlignment('center');
    sh.getRange(HR + 1, 11, n, 1).setFontWeight('bold');
    sh.getRange(HR, 1, n + 1, NC).createFilter();
  } else {
    sh.getRange(HR + 1, 1, 1, NC).merge().setValue('Belum ada data alat.')
      .setHorizontalAlignment('center').setFontColor('#55648a');
  }

  [40, 200, 110, 110, 120, 120, 120, 110, 100, 110, 120, 80, 90, 320, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  ss.setActiveSheet(sh);
}

function tes() {
  Logger.log(doGet({ parameter: {} }).getContent());
}
