// ============================================================
// LIST ALL NAMED RANGES — RaListAllNamedRanges.gs
// CUMA BACA — tidak mengubah apa-apa.
// Tujuan: melihat semua named range yang sekarang ada di seluruh
// spreadsheet (nama, sheet tempatnya, kolom keberapa), supaya kita
// bisa cek apakah ada korban lain selain RA_ID dari bug kemarin
// (baris yang menghapus SEMUA named range berawalan "RA_" secara
// tidak sengaja).
// ============================================================

function listAllNamedRanges() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ranges = ss.getNamedRanges();

  Logger.log('=== TOTAL NAMED RANGE DI SPREADSHEET INI: ' + ranges.length + ' ===');

  // Kelompokkan per sheet biar gampang dibaca
  const bySheet = {};
  ranges.forEach(function(nr) {
    const sheetName = nr.getRange().getSheet().getName();
    if (!bySheet[sheetName]) bySheet[sheetName] = [];
    bySheet[sheetName].push({ name: nr.getName(), col: nr.getRange().getColumn() });
  });

  Object.keys(bySheet).sort().forEach(function(sheetName) {
    Logger.log('--- Sheet: "' + sheetName + '" (' + bySheet[sheetName].length + ' named range) ---');
    bySheet[sheetName]
      .sort(function(a, b) { return a.col - b.col; })
      .forEach(function(r) {
        Logger.log('  Kolom ' + r.col + ': ' + r.name);
      });
  });

  Logger.log('=== SELESAI ===');
}
