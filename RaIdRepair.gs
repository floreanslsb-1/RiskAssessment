// ============================================================
// VERIFY & REPAIR RA_ID — RaIdRepair.gs
// CUMA BACA dulu di verifyRaIdColumn() — tidak mengubah apa-apa.
// Setelah yakin, baru jalankan repairRaIdNamedRange().
// ============================================================

function verifyRaIdColumn() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('2026');
  if (!sheet) { Logger.log('Sheet "2026" tidak ditemukan.'); return; }

  const lastRow = sheet.getLastRow();
  const sampleEnd = Math.min(lastRow, 20); // baris 13-20 (data mulai row 13)
  Logger.log('=== Isi kolom 12-16, baris 13-' + sampleEnd + ' (sheet "2026") ===');

  for (let r = 13; r <= sampleEnd; r++) {
    const vals = sheet.getRange(r, 12, 1, 5).getValues()[0]; // kolom 12..16
    Logger.log('Baris ' + r + ': ' +
      'Kol12=' + vals[0] + ' | Kol13(Link_Kajian_Risiko)=' + vals[1] +
      ' | Kol14(?)=' + vals[2] + ' | Kol15(Status_Kajian_Risiko)=' + vals[3] +
      ' | Kol16=' + vals[4]);
  }
}

function repairRaIdNamedRange() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('2026');
  if (!sheet) throw new Error('Sheet "2026" tidak ditemukan.');

  ss.setNamedRange('RA_ID', sheet.getRange(1, 14));
  CacheService.getScriptCache().remove('COL_MAP');
  Logger.log('=== SELESAI. Named range RA_ID dibuat di kolom 14, sheet "2026". Cache COL_MAP dibersihkan. ===');
}
