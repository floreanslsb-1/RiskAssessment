// ============================================================
// RA NAMED RANGE REPAIR — RaNamedRangeRepair.gs
// DARURAT: setupRaConfig() versi lama sempat menghapus SEMUA named
// range berawalan "RA_" di seluruh spreadsheet, termasuk named range
// asli sheet Kajian_Risiko (RA_RA_ID, RA_STATUS, dst). Isi sel di
// sheet TIDAK hilang — cuma "penanda" nama kolomnya yang hilang,
// sehingga getRaColumnMap() gagal menemukannya.
//
// CARA PAKAI:
//   1. Jalankan diagnoseRaHeaders() dulu — ini CUMA BACA, tidak
//      mengubah apa-apa. Cek log-nya (View → Logs / Ctrl+Enter),
//      pastikan tiap label ketemu kolomnya dengan masuk akal.
//   2. Kalau hasilnya sudah benar, jalankan repairRaNamedRanges() —
//      ini yang benar-benar membuat ulang named range-nya.
// ============================================================

const RA_HEADER_TO_RANGE_NAME = {
  'RA_ID':             'RA_RA_ID',
  'Status':             'RA_STATUS',
  'Saved_At':           'RA_SAVED_AT',
  'Is_Submitted':       'RA_IS_SUBMITTED',
  'Working_Area':       'RA_WORKING_AREA',
  'Job_Title':          'RA_JOB_TITLE',
  'Workstation':        'RA_WORKSTATION',
  'Aktivitas':          'RA_AKTIVITAS',
  'Rutin':              'RA_RUTIN',
  'Kategori_Risiko':    'RA_KATEGORI_RISIKO',
  'Asumsi_Bahaya':      'RA_ASUMSI_BAHAYA',
  'Dampak':             'RA_DAMPAK',
  'Jml_Terpapar':       'RA_JML_TERPAPAR',
  'Ibu_Hamil':          'RA_IBU_HAMIL',
  'Ibu_Menyusui':       'RA_IBU_MENYUSUI',
  'Penyakit_Khusus':    'RA_PENYAKIT_KHUSUS',
  'Disabilitas':        'RA_DISABILITAS',
  'ROP_Before':         'RA_ROP_BEFORE',
  'RS_Before':          'RA_RS_BEFORE',
  'Score_Before':       'RA_SCORE_BEFORE',
  'Level_Before':       'RA_LEVEL_BEFORE',
  'Risk_Controls_JSON': 'RA_RISK_CONTROLS',
  'ROP_After':          'RA_ROP_AFTER',
  'RS_After':           'RA_RS_AFTER',
  'Score_After':        'RA_SCORE_AFTER',
  'Level_After':        'RA_LEVEL_AFTER',
  'Dept':               'RA_DEPT',
};

function diagnoseRaHeaders() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
  if (!sheet) { Logger.log('Sheet ' + CONFIG.RA_SHEET_NAME + ' tidak ditemukan.'); return; }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  Logger.log('=== ISI ROW 1 (header) sheet "' + CONFIG.RA_SHEET_NAME + '" ===');
  headers.forEach(function(h, i) {
    Logger.log('Kolom ' + (i + 1) + ': "' + h + '"');
  });

  Logger.log('=== PENCOCOKAN dengan named range yang diharapkan ===');
  Object.keys(RA_HEADER_TO_RANGE_NAME).forEach(function(label) {
    const colIndex  = headers.indexOf(label);
    const rangeName = RA_HEADER_TO_RANGE_NAME[label];
    if (colIndex === -1) {
      Logger.log('❌ TIDAK KETEMU: label "' + label + '" (harusnya jadi ' + rangeName + ')');
    } else {
      Logger.log('✅ "' + label + '" ditemukan di kolom ' + (colIndex + 1) + ' → akan jadi ' + rangeName);
    }
  });
}

function repairRaNamedRanges() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
  if (!sheet) throw new Error('Sheet ' + CONFIG.RA_SHEET_NAME + ' tidak ditemukan.');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let created = 0, skipped = 0;
  Object.keys(RA_HEADER_TO_RANGE_NAME).forEach(function(label) {
    const colIndex  = headers.indexOf(label);
    const rangeName = RA_HEADER_TO_RANGE_NAME[label];
    if (colIndex === -1) {
      Logger.log('⚠ Lewati ' + rangeName + ' — label "' + label + '" tidak ketemu di row 1.');
      skipped++;
      return;
    }
    ss.setNamedRange(rangeName, sheet.getRange(1, colIndex + 1));
    Logger.log('Dibuat: ' + rangeName + ' → kolom ' + (colIndex + 1));
    created++;
  });

  CacheService.getScriptCache().remove('RA_COL_MAP');
  Logger.log('=== SELESAI. Dibuat: ' + created + ', dilewati: ' + skipped + ' ===');
}
