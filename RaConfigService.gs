// ============================================================
// RA CONFIG SERVICE — RaConfigService.gs
// Membaca data dari sheet "RA_Config" lewat named range marker,
// dengan caching (pola sama seperti getColumnMap()/getRaColumnMap()
// yang sudah ada di Code.gs).
// ============================================================

function getRaConfig() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('RA_CONFIG');
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  const docNo   = ss.getRangeByName('RA_DOC_NO').getValue();
  const docRev  = ss.getRangeByName('RA_DOC_REV').getValue();
  const effDate = ss.getRangeByName('RA_DOC_EFFECTIVE_DATE').getValue();
  const mainUrl = ss.getRangeByName('RA_MAIN_APP_URL').getValue();

  const kategoriList = readTableBelow_(ss, 'RA_KATEGORI_TABLE')
    .map(function(r) { return r[0]; })
    .filter(String);

  const ropRows   = readTableBelow_(ss, 'RA_ROP_LABEL_TABLE');
  const ropLabels = {};
  ropRows.forEach(function(r) { if (r[0] !== '') ropLabels[r[0]] = r[1]; });

  const rsRows = readTableBelow_(ss, 'RA_RS_LABEL_TABLE');
  const rsLabelsByKategori = {};
  rsRows.forEach(function(r) {
    const kat = r[0], lvl = r[1], label = r[2];
    if (!kat) return;
    if (!rsLabelsByKategori[kat]) rsLabelsByKategori[kat] = {};
    rsLabelsByKategori[kat][lvl] = label;
  });

  const asumsiRows = readTableBelow_(ss, 'RA_ASUMSI_TABLE');
  const asumsiByKategori = {};
  asumsiRows.forEach(function(r) {
    const kat = r[0], asumsi = r[1];
    if (!kat) return;
    if (!asumsiByKategori[kat]) asumsiByKategori[kat] = [];
    asumsiByKategori[kat].push(asumsi);
  });

  const config = {
    docNo:                  docNo,
    docRev:                 docRev,
    effectiveDate:          effDate,
    effectiveDateFormatted: effDate ? Utilities.formatDate(effDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
    mainAppUrl:             mainUrl,
    kategoriList:       kategoriList,
    ropLabels:          ropLabels,
    rsLabelsByKategori: rsLabelsByKategori,
    asumsiByKategori:   asumsiByKategori,
  };

  cache.put('RA_CONFIG', JSON.stringify(config), 21600); // 6 jam, sama seperti COL_MAP
  return config;
}

// Baca semua baris di bawah named range penanda (header row), berhenti
// begitu kolom pertama kosong. Tidak butuh update range kalau baris
// ditambah/dihapus dari Admin Panel nantinya.
function readTableBelow_(ss, markerName) {
  const marker = ss.getRangeByName(markerName);
  if (!marker) throw new Error('Named range tidak ditemukan: ' + markerName);
  const sheet    = marker.getSheet();
  const startRow = marker.getRow() + 1;
  const startCol = marker.getColumn();
  const numCols  = marker.getNumColumns();
  const lastRow  = sheet.getLastRow();
  if (lastRow < startRow) return [];

  const values = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols).getValues();
  const result = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) break;
    result.push(values[i]);
  }
  return result;
}

function clearRaConfigCache() {
  CacheService.getScriptCache().remove('RA_CONFIG');
}
