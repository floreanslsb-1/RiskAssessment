// ============================================================
// RISK ASSESSMENT WEB APP — Code.gs  [PATCHED]
// PT. Sayap Mas Utama | Form No: SMU/IMS-00/06-013 | Rev: 01
// Fixes applied:
//   #1  writeRaSheetHeaders_ used hardcoded 27 columns → now dynamic
//   #2  generateRaExcel leaked temp spreadsheet on error → fixed with try/finally
//   #3  buildExcelSheet_ header collision at row 5 → headers moved to row 7, data row 8+
//   #5  autoSaveRa had no LockService → added
//   #6  submitRa had no LockService → added
//   #7  findRaRow_ called inside loop (N sheet reads) → read once outside loop
// ============================================================

const CONFIG = {
  SHEET_ID:        '1-dVtUrAJn6Yvo3KzJJemmbgn2pR7-APM7I5KX36tb6c',
  RA_SHEET_NAME:   'Kajian_Risiko',
  MAIN_SHEET_NAME: 'Template', // fallback lawas saja — pencarian utama sekarang lewat sheet tahun (lihat findMainSheetRowByRaId_)
  MAIN_APP_URL:    'https://script.google.com/a/macros/wingscorp.com/s/AKfycbxZxJh54v-YO8zQ0TFBg08yzrFHJL9tCXBPmKkC1fcbOCx3wx431Bz86zyPNVATszJP/exec',
  LOGO_FILE_ID:    '1gVRF0XPZ806GOF58AK7qI7DaY6VMopdl',
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.admin === '1') {
    if (!isRaAdmin_()) {
      return HtmlService.createHtmlOutput(
        '<p style="font-family:sans-serif;padding:40px;color:#6B7280">Akses ditolak. Halaman ini khusus admin RA app.</p>'
      );
    }
    const adminTemplate = HtmlService.createTemplateFromFile('RaAdmin');
    adminTemplate.raConfig    = getRaConfig();
    adminTemplate.isAdmin     = true;
    adminTemplate.adminEmails = getRaAdminEmailsList_();
    return adminTemplate.evaluate()
      .setTitle('Admin Panel — Kajian Risiko')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const raId      = params.raId      || '';
  const returnUrl = params.returnUrl || CONFIG.MAIN_APP_URL;
  const raConfig  = getRaConfig();
  const selfUrl   = ScriptApp.getService().getUrl();

  // Info FUP No./Dept ini cuma pemanis tampilan — kalau gagal diambil karena
  // apapun (struktur sheet MOC Portal berubah, dst), form RA tetap harus bisa
  // dibuka dan diisi, bukan malah blank error.
  let fupNo = '', fupDept = '';
  try {
    const fupResult = raId ? findMainSheetRowByRaId_(raId) : null;
    if (fupResult) {
      const col = getColumnMap();
      fupNo   = fupResult.rowData[col.No_Registrasi - 1] || '';
      fupDept = fupResult.rowData[col.Departemen   - 1] || '';
    }
  } catch (err) {
    console.error('doGet: gagal ambil info FUP — ' + err.message);
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.raId      = raId;
  template.returnUrl = returnUrl;
  template.logoB64   = getLogoBase64_();
  template.fupNo     = fupNo;
  template.fupLink   = returnUrl;
  template.fupDept   = fupDept;
  template.raConfig  = raConfig;
  template.selfUrl   = selfUrl;

  return template.evaluate()
    .setTitle('Kajian Risiko — PT. Sayap Mas Utama')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getLogoBase64_() {
  try {
    if (CONFIG.LOGO_FILE_ID && CONFIG.LOGO_FILE_ID !== 'YOUR_LOGO_FILE_ID_IN_DRIVE') {
      const blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
      return 'data:image/png;base64,' + Utilities.base64Encode(blob.getBytes());
    }
  } catch(e) {}
  return '';
}

function loadRaData(raId) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Sheet not found', data: null };

    const dept = getDeptFromMainSheet_(raId);
    if (!raId) return { success: true, data: null, dept: dept };

    const col      = getRaColumnMap();
    const data     = sheet.getDataRange().getValues();
    const baseRaId = raId.replace(/_\d+$/, '');
    Logger.log('DEBUG loadRaData: raId=' + raId + ' baseRaId=' + baseRaId + ' totalDataRows=' + data.length + ' col.RA_RA_ID=' + col['RA_RA_ID'] + ' dept=' + dept);

    const matchedRows = [];
    for (let i = 1; i < data.length; i++) {
      const rowRaId     = String(data[i][col['RA_RA_ID'] - 1]);
      const rowBaseRaId = rowRaId.replace(/_\d+$/, '');
      if (rowBaseRaId === baseRaId) {
        matchedRows.push({ sheetRowIndex: i + 1, data: data[i] });
      }
    }

    if (matchedRows.length === 0) return { success: true, data: null, dept: dept };

    const allRisks = matchedRows.map(function(r) {
      const row      = r.data;
      const controls = row[col['RA_RISK_CONTROLS'] - 1];
      return {
        raId:           row[col['RA_RA_ID']          - 1],
        savedAt:        row[col['RA_SAVED_AT']       - 1],
        status:         row[col['RA_STATUS']         - 1],
        workingArea:    row[col['RA_WORKING_AREA']   - 1],
        jobTitle:       row[col['RA_JOB_TITLE']      - 1],
        workstation:    row[col['RA_WORKSTATION']    - 1],
        aktivitas:      row[col['RA_AKTIVITAS']      - 1],
        rutin:          row[col['RA_RUTIN']          - 1],
        kategoriRisiko: row[col['RA_KATEGORI_RISIKO']- 1],
        assumsiBahaya:  row[col['RA_ASUMSI_BAHAYA']  - 1],
        dampak:         row[col['RA_DAMPAK']         - 1],
        jmlTerpapar:    row[col['RA_JML_TERPAPAR']   - 1],
        ibuHamil:       row[col['RA_IBU_HAMIL']      - 1],
        ibuMenyusui:    row[col['RA_IBU_MENYUSUI']   - 1],
        penyakitKhusus: row[col['RA_PENYAKIT_KHUSUS']- 1],
        disabilitas:    row[col['RA_DISABILITAS']    - 1],
        ropBefore:      row[col['RA_ROP_BEFORE']     - 1],
        rsBefore:       row[col['RA_RS_BEFORE']      - 1],
        scoreBefore:    row[col['RA_SCORE_BEFORE']   - 1],
        levelBefore:    row[col['RA_LEVEL_BEFORE']   - 1],
        riskControls:   controls ? JSON.parse(controls) : [],
        ropAfter:       row[col['RA_ROP_AFTER']      - 1],
        rsAfter:        row[col['RA_RS_AFTER']       - 1],
        scoreAfter:     row[col['RA_SCORE_AFTER']    - 1],
        levelAfter:     row[col['RA_LEVEL_AFTER']    - 1],
        isSubmitted:    row[col['RA_IS_SUBMITTED']   - 1],
        dept:           row[col['RA_DEPT']           - 1],
        _sheetRowIndex: r.sheetRowIndex,
      };
    });

    return { success: true, dept: dept, allRisks: allRisks };
  } catch(err) {
    Logger.log('DEBUG loadRaData ERROR: ' + err.message + ' | stack: ' + err.stack);
    return { success: false, error: err.message, data: null };
  }
}

// ── AUTOSAVE ─────────────────────────────────────────────────
// FIX #5: Added LockService — prevents parallel writes from rapid autosave
// FIX #7: sheetData read once outside the forEach loop
function autoSaveRa(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet   = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.RA_SHEET_NAME);
      writeRaSheetHeaders_(sheet);
    }

    const now      = new Date();
    const allRisks = payload.allRisks && payload.allRisks.length > 0
                     ? payload.allRisks : [payload];
    const baseRaId = payload.raId || '';
    if (!baseRaId) return { success: false, error: 'RA ID missing' };

    const col       = getRaColumnMap();
    const sheetData = sheet.getDataRange().getValues(); // FIX #7: read once
    Logger.log('DEBUG autoSaveRa: baseRaId=' + baseRaId + ' allRisksCount=' + allRisks.length + ' col.RA_RA_ID=' + col['RA_RA_ID']);

    allRisks.forEach(function(r, i) {
      const rowRaId  = baseRaId + (i === 0 ? '' : '_' + i);
      const rowIndex = findRaRowFromData_(sheetData, rowRaId); // FIX #7: use pre-read data
      const rPayload = Object.assign({}, payload, r, { raId: rowRaId, baseRaId: baseRaId });
      const status    = computeStatus_(r.riskControls || [], r.assumsiBahaya || '');
      const evalAfter = computeEvalAfter_(r.ropBefore, r.rsBefore, r.riskControls || []);
      const rowData   = buildRowData_(rPayload, status, evalAfter, now, false, col);
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        const lastRow = findLastRaRowFromData_(sheetData, baseRaId); // FIX #7
        if (lastRow > 0) {
          sheet.insertRowAfter(lastRow);
          sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
        } else {
          sheet.appendRow(rowData);
        }
      }
    });

    updateMainSheetKajianStatus_(baseRaId, 'DRAFT');
    return { success: true, savedAt: now.toISOString(), raId: baseRaId };
  } catch(err) {
    Logger.log('DEBUG autoSaveRa ERROR: ' + err.message + ' | stack: ' + err.stack);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ── SUBMIT (FINAL) ───────────────────────────────────────────
// FIX #6: Added LockService — prevents double-submit race condition
// FIX #7: sheetData read once outside the forEach loop
function submitRa(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet   = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.RA_SHEET_NAME);
      writeRaSheetHeaders_(sheet);
    }

    const now      = new Date();
    const allRisks = payload.allRisks && payload.allRisks.length > 0
                     ? payload.allRisks : [payload];
    const baseRaId = payload.raId || '';
    if (!baseRaId) return { success: false, error: 'RA ID missing' };

    const col       = getRaColumnMap();
    const sheetData = sheet.getDataRange().getValues(); // FIX #7: read once

    allRisks.forEach(function(r, i) {
      const rowRaId   = baseRaId + (i === 0 ? '' : '_' + i);
      const rowIndex  = findRaRowFromData_(sheetData, rowRaId); // FIX #7
      const rPayload  = Object.assign({}, payload, r, { raId: rowRaId, baseRaId: baseRaId });
      const status    = computeStatus_(r.riskControls || [], r.assumsiBahaya || '');
      const evalAfter = computeEvalAfter_(r.ropBefore, r.rsBefore, r.riskControls || []);
      const rowData   = buildRowData_(rPayload, status, evalAfter, now, true, col);
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        const lastRow = findLastRaRowFromData_(sheetData, baseRaId); // FIX #7
        if (lastRow > 0) {
          sheet.insertRowAfter(lastRow);
          sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
        } else {
          sheet.appendRow(rowData);
        }
      }
    });

    updateMainSheetKajianStatus_(baseRaId, 'SELESAI');
    return { success: true };
  } catch(err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ── HELPERS ──────────────────────────────────────────────────
function findMainSheetRowByRaId_(raId) {
  try {
    if (!raId) return null;
    const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const col      = getColumnMap();
    const baseRaId = raId.replace(/_\d+$/, '');

    // MOC Portal sekarang punya satu sheet per tahun (mis. "2026", "2027").
    // Dokumen yang lagi diisi Risk Assessment-nya bisa saja dibuat tahun ini
    // atau tahun lalu, jadi kita coba beberapa sheet tahun sebelum menyerah.
    // Fallback ke CONFIG.MAIN_SHEET_NAME kalau suatu saat MOC Portal masih
    // pakai skema satu-sheet lama.
    const currentYear = new Date().getFullYear();
    const candidates  = [];
    for (let y = currentYear; y >= currentYear - 2; y--) {
      const s = ss.getSheetByName(String(y));
      if (s) candidates.push(s);
    }
    const legacy = ss.getSheetByName(CONFIG.MAIN_SHEET_NAME);
    if (legacy && candidates.indexOf(legacy) === -1) candidates.push(legacy);

    for (const sheet of candidates) {
      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        const cellRaId = String(data[i][col.RA_ID - 1] || '').replace(/_\d+$/, '');
        if (cellRaId && cellRaId === baseRaId) {
          return { rowNumber: i + 1, rowData: data[i], sheet: sheet };
        }
      }
    }
    return null;
  } catch(e) { return null; }
}

// Original version (kept for callers outside save loops)
function findRaRow_(sheet, raId) {
  return findRaRowFromData_(sheet.getDataRange().getValues(), raId);
}

// FIX #7: Accepts pre-read data array — avoids repeated sheet reads inside loops
function findRaRowFromData_(data, raId) {
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(raId)) return i + 1;
  }
  return -1;
}

function findLastRaRow_(sheet, baseRaId) {
  return findLastRaRowFromData_(sheet.getDataRange().getValues(), baseRaId);
}

// FIX #7: Accepts pre-read data array
function findLastRaRowFromData_(data, baseRaId) {
  let lastRow = -1;
  for (let i = 1; i < data.length; i++) {
    const rowBaseRaId = String(data[i][0]).replace(/_\d+$/, '');
    if (rowBaseRaId === baseRaId) lastRow = i + 1;
  }
  return lastRow;
}

function buildRowData_(payload, status, evalAfter, now, isSubmitted, col) {
  const totalCols = Math.max(
    col['RA_RA_ID'],         col['RA_STATUS'],          col['RA_SAVED_AT'],
    col['RA_IS_SUBMITTED'],  col['RA_WORKING_AREA'],    col['RA_JOB_TITLE'],
    col['RA_WORKSTATION'],   col['RA_AKTIVITAS'],       col['RA_RUTIN'],
    col['RA_KATEGORI_RISIKO'],col['RA_ASUMSI_BAHAYA'],  col['RA_DAMPAK'],
    col['RA_JML_TERPAPAR'],  col['RA_IBU_HAMIL'],       col['RA_IBU_MENYUSUI'],
    col['RA_PENYAKIT_KHUSUS'],col['RA_DISABILITAS'],    col['RA_ROP_BEFORE'],
    col['RA_RS_BEFORE'],     col['RA_SCORE_BEFORE'],    col['RA_LEVEL_BEFORE'],
    col['RA_RISK_CONTROLS'], col['RA_ROP_AFTER'],       col['RA_RS_AFTER'],
    col['RA_SCORE_AFTER'],   col['RA_LEVEL_AFTER'],     col['RA_DEPT']
  );

  const row = new Array(totalCols).fill('');
  row[col['RA_RA_ID']          - 1] = payload.raId            || '';
  row[col['RA_STATUS']         - 1] = status                  || '';
  row[col['RA_SAVED_AT']       - 1] = now;
  row[col['RA_IS_SUBMITTED']   - 1] = isSubmitted;
  row[col['RA_WORKING_AREA']   - 1] = payload.workingArea     || '';
  row[col['RA_JOB_TITLE']      - 1] = payload.jobTitle        || '';
  row[col['RA_WORKSTATION']    - 1] = payload.workstation     || '';
  row[col['RA_AKTIVITAS']      - 1] = payload.aktivitas       || '';
  row[col['RA_RUTIN']          - 1] = payload.rutin           || '';
  row[col['RA_KATEGORI_RISIKO']- 1] = payload.kategoriRisiko  || '';
  row[col['RA_ASUMSI_BAHAYA']  - 1] = payload.assumsiBahaya   || '';
  row[col['RA_DAMPAK']         - 1] = payload.dampak          || '';
  row[col['RA_JML_TERPAPAR']   - 1] = payload.jmlTerpapar     || '';
  row[col['RA_IBU_HAMIL']      - 1] = payload.ibuHamil        || '';
  row[col['RA_IBU_MENYUSUI']   - 1] = payload.ibuMenyusui     || '';
  row[col['RA_PENYAKIT_KHUSUS']- 1] = payload.penyakitKhusus  || '';
  row[col['RA_DISABILITAS']    - 1] = payload.disabilitas     || '';
  row[col['RA_ROP_BEFORE']     - 1] = payload.ropBefore       || '';
  row[col['RA_RS_BEFORE']      - 1] = payload.rsBefore        || '';
  row[col['RA_SCORE_BEFORE']   - 1] = payload.scoreBefore     || '';
  row[col['RA_LEVEL_BEFORE']   - 1] = payload.levelBefore     || '';
  row[col['RA_RISK_CONTROLS']  - 1] = JSON.stringify(payload.riskControls || []);
  row[col['RA_ROP_AFTER']      - 1] = evalAfter.ropAfter      || '';
  row[col['RA_RS_AFTER']       - 1] = evalAfter.rsAfter       || '';
  row[col['RA_SCORE_AFTER']    - 1] = evalAfter.scoreAfter    || '';
  row[col['RA_LEVEL_AFTER']    - 1] = evalAfter.levelAfter    || '';
  row[col['RA_DEPT']           - 1] = payload.dept            || '';
  return row;
}

// ── RISK CALCULATION ─────────────────────────────────────────
const ROP_LEVELS = [10, 8, 6, 4, 1];
const RS_LEVELS  = [100, 40, 21, 8, 2];

function computeEvalAfter_(ropBefore, rsBefore, controls) {
  let ropIdx = ROP_LEVELS.indexOf(Number(ropBefore));
  let rsIdx  = RS_LEVELS.indexOf(Number(rsBefore));
  if (ropIdx < 0) ropIdx = 0;
  if (rsIdx  < 0) rsIdx  = 0;

  let hasAdminControl = false, hasPPE = false;

  controls.forEach(function(c) {
    const strat = (c.strategi || '').toLowerCase();
    const kat   = (c.kategori || '').toLowerCase();
    const done  = c.status === 'Done'; // FIX: konsisten dengan client — skor cuma turun kalau Done
    if (strat === 'mitigate' && kat === 'engineering control' && done) {
      ropIdx = Math.min(ropIdx + 1, ROP_LEVELS.length - 1);
      rsIdx  = Math.min(rsIdx  + 1, RS_LEVELS.length  - 1);
    }
    if ((strat === 'mitigate' || strat === 'acceptance') && kat === 'administration control' && done) {
      if (!hasAdminControl) { ropIdx = Math.min(ropIdx + 1, ROP_LEVELS.length - 1); hasAdminControl = true; }
    }
    if (strat === 'acceptance' && kat === 'ppe / apd' && done) {
      if (!hasPPE) { rsIdx = Math.min(rsIdx + 1, RS_LEVELS.length - 1); hasPPE = true; }
    }
  });

  const ropAfter = ROP_LEVELS[ropIdx], rsAfter = RS_LEVELS[rsIdx];
  return { ropAfter, rsAfter, scoreAfter: ropAfter * rsAfter, levelAfter: computeRiskLevel_(ropAfter * rsAfter) };
}

function computeRiskLevel_(score) {
  if (score >= 210) return 'High';
  if (score >= 20)  return 'Medium';
  return 'Low';
}

function computeStatus_(controls, assumsiBahaya) {
  if (!assumsiBahaya) return 'Active';
  let hasEliminasi = false, hasTransfer = false, hasSub = false;
  controls.forEach(function(c) {
    const strat = (c.strategi || '').toLowerCase();
    const kat   = (c.kategori || '').toLowerCase();
    if (strat === 'avoidance' && kat === 'eliminasi') hasEliminasi = true;
    if (strat === 'transference')                      hasTransfer  = true;
    if (strat === 'mitigate' && kat === 'subtitusi')  hasSub       = true;
  });
  if (hasEliminasi) return 'Eliminasi';
  if (hasTransfer)  return 'Transference';
  if (hasSub)       return 'Substitute';
  return 'Active';
}

// FIX #1: Dynamic column count — no longer hardcoded to 27
function writeRaSheetHeaders_(sheet) {
  const col = getRaColumnMap();
  const totalCols = Math.max(
    col['RA_RA_ID'],         col['RA_STATUS'],          col['RA_SAVED_AT'],
    col['RA_IS_SUBMITTED'],  col['RA_WORKING_AREA'],    col['RA_JOB_TITLE'],
    col['RA_WORKSTATION'],   col['RA_AKTIVITAS'],       col['RA_RUTIN'],
    col['RA_KATEGORI_RISIKO'],col['RA_ASUMSI_BAHAYA'],  col['RA_DAMPAK'],
    col['RA_JML_TERPAPAR'],  col['RA_IBU_HAMIL'],       col['RA_IBU_MENYUSUI'],
    col['RA_PENYAKIT_KHUSUS'],col['RA_DISABILITAS'],    col['RA_ROP_BEFORE'],
    col['RA_RS_BEFORE'],     col['RA_SCORE_BEFORE'],    col['RA_LEVEL_BEFORE'],
    col['RA_RISK_CONTROLS'], col['RA_ROP_AFTER'],       col['RA_RS_AFTER'],
    col['RA_SCORE_AFTER'],   col['RA_LEVEL_AFTER'],     col['RA_DEPT']
  );

  const headers = new Array(totalCols).fill('');
  headers[col['RA_RA_ID']          - 1] = 'RA_ID';
  headers[col['RA_STATUS']         - 1] = 'Status';
  headers[col['RA_SAVED_AT']       - 1] = 'Saved_At';
  headers[col['RA_IS_SUBMITTED']   - 1] = 'Is_Submitted';
  headers[col['RA_WORKING_AREA']   - 1] = 'Working_Area';
  headers[col['RA_JOB_TITLE']      - 1] = 'Job_Title';
  headers[col['RA_WORKSTATION']    - 1] = 'Workstation';
  headers[col['RA_AKTIVITAS']      - 1] = 'Aktivitas';
  headers[col['RA_RUTIN']          - 1] = 'Rutin';
  headers[col['RA_KATEGORI_RISIKO']- 1] = 'Kategori_Risiko';
  headers[col['RA_ASUMSI_BAHAYA']  - 1] = 'Asumsi_Bahaya';
  headers[col['RA_DAMPAK']         - 1] = 'Dampak';
  headers[col['RA_JML_TERPAPAR']   - 1] = 'Jml_Terpapar';
  headers[col['RA_IBU_HAMIL']      - 1] = 'Ibu_Hamil';
  headers[col['RA_IBU_MENYUSUI']   - 1] = 'Ibu_Menyusui';
  headers[col['RA_PENYAKIT_KHUSUS']- 1] = 'Penyakit_Khusus';
  headers[col['RA_DISABILITAS']    - 1] = 'Disabilitas';
  headers[col['RA_ROP_BEFORE']     - 1] = 'ROP_Before';
  headers[col['RA_RS_BEFORE']      - 1] = 'RS_Before';
  headers[col['RA_SCORE_BEFORE']   - 1] = 'Score_Before';
  headers[col['RA_LEVEL_BEFORE']   - 1] = 'Level_Before';
  headers[col['RA_RISK_CONTROLS']  - 1] = 'Risk_Controls_JSON';
  headers[col['RA_ROP_AFTER']      - 1] = 'ROP_After';
  headers[col['RA_RS_AFTER']       - 1] = 'RS_After';
  headers[col['RA_SCORE_AFTER']    - 1] = 'Score_After';
  headers[col['RA_LEVEL_AFTER']    - 1] = 'Level_After';
  headers[col['RA_DEPT']           - 1] = 'Dept';

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// FIX #2: try/finally ensures temp spreadsheet is always cleaned up
// FIX #3: Row layout — metadata rows 1-6, headers row 7, data row 8+
function generateRaExcel(payload) {
  let tempId = null;
  try {
    const data = {
      dept:        payload.dept    || '',
      savedAt:     new Date().toISOString(),
      fupNo:       payload.fupNo   || '',
      fupLink:     payload.fupLink || '',
      isSubmitted: payload.isSubmitted || false,
      allRisks: (payload.allRisks || [payload]).map(function(r) {
        const after = computeEvalAfter_(r.ropBefore, r.rsBefore, r.riskControls || []);
        return Object.assign({}, r, {
          status:     computeStatus_(r.riskControls || [], r.assumsiBahaya || ''),
          ropAfter:   after.ropAfter, rsAfter: after.rsAfter,
          scoreAfter: after.scoreAfter, levelAfter: after.levelAfter,
        });
      }),
    };

    const ss    = SpreadsheetApp.create('RA_temp_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    tempId      = ss.getId(); // saved for finally cleanup
    buildExcelSheet_(sheet, data);

    const url  = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob();
    return Utilities.base64Encode(blob.getBytes());
  } catch(err) {
    throw new Error('generateRaExcel: ' + err.message);
  } finally {
    // FIX #2: Always clean up temp spreadsheet — even if fetch threw
    if (tempId) {
      try { DriveApp.getFileById(tempId).setTrashed(true); } catch(e) {}
    }
  }
}

// FIX #3: Fixed row collision — layout is now:
//   Rows 1-4: company/dept/date/FUP metadata
//   Row  5:   (empty — spacer between meta and status)
//   Row  6:   DRAFT / FINAL status
//   Row  7:   Column headers
//   Row  8+:  Data rows
//   Sig rows: after data
function buildExcelSheet_(sheet, data) {
  const risks = data.allRisks || [];
  const raConfig = getRaConfig();

  sheet.getRange(1, 1).setValue('FORM RISK ASSESSMENT (RA)');
  sheet.getRange(2, 1).setValue('Dept: ' + (data.dept || ''));
  sheet.getRange(2, 5).setValue('No. Dokumen: ' + raConfig.docNo + '  |  Revisi: ' + raConfig.docRev);
  sheet.getRange(3, 1).setValue('Tgl. Update: ' + (data.savedAt || '').slice(0, 10));
  sheet.getRange(4, 1).setValue('No. FUP: ' + (data.fupNo || '-'));
  sheet.getRange(4, 5).setValue('Link FUP: ' + (data.fupLink || '-'));
  // Row 5 left empty as visual spacer
  // Row 6: DRAFT/FINAL status (FIX #3: was row 5, collided with headers)
  if (!data.isSubmitted) {
    sheet.getRange(6, 1).setValue('⚠ DRAFT — Belum disubmit');
    sheet.getRange(6, 1).setFontColor('#DC2626').setFontWeight('bold');
  } else {
    sheet.getRange(6, 1).setValue('✓ FINAL — Sudah disubmit');
    sheet.getRange(6, 1).setFontColor('#059669').setFontWeight('bold');
  }

  // Row 7: Column headers (FIX #3: was row 5)
  const headers = [
    'No','Status','Working Area','Job Title','Workstation','Aktivitas','R/NR',
    'Kategori','Asumsi/Bahaya','Dampak','Jml Terpapar',
    'Ibu Hamil','Ibu Menyusui','Penyakit Khusus','Disabilitas',
    'ROP Before','RS Before','Score Before','Level Before',
    'ENG Control','ADM Control','PPE/APD',
    'ROP After','RS After','Score After','Level After'
  ];
  const maxCtrl = risks.length > 0 ? Math.max(...risks.map(r => (r.riskControls||[]).length), 1) : 1;
  for (let i = 0; i < maxCtrl; i++) {
    headers.push('Strategi '+(i+1),'Kategori '+(i+1),'Activity '+(i+1),'PIC '+(i+1),'Due Date '+(i+1),'Status '+(i+1));
  }
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]);

  // Rows 8+: Data (FIX #3: was row 6)
  risks.forEach(function(r, idx) {
    const rop_b = Number(r.ropBefore)||0, rs_b = Number(r.rsBefore)||0;
    const score_b = rop_b * rs_b;
    const level_b = score_b ? (score_b >= 210 ? 'HIGH' : score_b >= 20 ? 'MEDIUM' : 'LOW') : '';
    const rop_a = Number(r.ropAfter)||0, rs_a = Number(r.rsAfter)||0;
    const score_a = rop_a * rs_a;
    const level_a = score_a ? (score_a >= 210 ? 'HIGH' : score_a >= 20 ? 'MEDIUM' : 'LOW') : '';
    const isOH  = r.kategoriRisiko === 'OH' || r.kategoriRisiko === 'OS Proses & Human';
    const ctrls = r.riskControls || [];
    const eng = ctrls.filter(c=>(c.kategori||'').toLowerCase()==='engineering control').map(c=>c.controlActivity||'').join(' | ');
    const adm = ctrls.filter(c=>(c.kategori||'').toLowerCase()==='administration control').map(c=>c.controlActivity||'').join(' | ');
    const ppe = ctrls.filter(c=>(c.kategori||'').toLowerCase()==='ppe / apd').map(c=>c.controlActivity||'').join(' | ');

    const rowData = [
      idx+1, r.status||'', r.workingArea||'', r.jobTitle||'', r.workstation||'',
      r.aktivitas||'', r.rutin||'', r.kategoriRisiko||'', r.assumsiBahaya||'',
      r.dampak||'', isOH?(r.jmlTerpapar||''):'—', isOH?(r.ibuHamil||''):'—',
      isOH?(r.ibuMenyusui||''):'—', isOH?(r.penyakitKhusus||''):'—',
      isOH?(r.disabilitas||''):'—',
      rop_b||'', rs_b||'', score_b||'', level_b,
      eng||'—', adm||'—', ppe||'—',
      rop_a||'', rs_a||'', score_a||'', level_a,
    ];
    ctrls.forEach(function(c) {
      rowData.push(c.strategi||'',c.kategori||'',c.controlActivity||'',c.pic||'',c.dueDate||'',c.status||'');
    });
    sheet.getRange(8 + idx, 1, 1, rowData.length).setValues([rowData]);
  });

  const sigRow = 8 + risks.length + 2;
  sheet.getRange(sigRow,   1).setValue('Dibuat');
  sheet.getRange(sigRow,  10).setValue('Disetujui');
  sheet.getRange(sigRow,  20).setValue('Diketahui');
  sheet.getRange(sigRow+2, 1).setValue('Section Head');
  sheet.getRange(sigRow+2,10).setValue('Department Head');
  sheet.getRange(sigRow+2,20).setValue('Management Representative');
}

function updateMainSheetKajianStatus_(raId, status) {
  try {
    const result = findMainSheetRowByRaId_(raId);
    if (!result || !result.sheet) { console.error('updateMainSheetKajianStatus_: row not found for raId: ' + raId); return; }
    const col = getColumnMap();
    result.sheet.getRange(result.rowNumber, col.Status_Kajian_Risiko).setValue(status);
  } catch(e) { console.error('updateMainSheetKajianStatus_ failed: ' + e.message); }
}

function getDeptFromMainSheet_(raId) {
  try {
    if (!raId) return '';
    const col    = getColumnMap();
    const result = findMainSheetRowByRaId_(raId);
    return result ? (result.rowData[col.Departemen - 1] || '') : '';
  } catch(e) { return ''; }
}

function getFupStatus(raId) {
  try {
    if (!raId) return '';
    const col    = getColumnMap();
    const result = findMainSheetRowByRaId_(raId);
    return result ? (result.rowData[col.Status - 1] || '') : '';
  } catch(e) { return ''; }
}

function getRaColumnMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RA_COL_MAP');
  if (cached) return JSON.parse(cached);
  const ss        = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const raSheet   = ss.getSheetByName(CONFIG.RA_SHEET_NAME);
  const raSheetId = raSheet.getSheetId();
  const map       = {};
  ss.getNamedRanges().forEach(nr => {
    if (nr.getRange().getSheet().getSheetId() === raSheetId)
      map[nr.getName().trim()] = nr.getRange().getColumn();
  });
  if (!map['RA_RA_ID']) throw new Error('RA named ranges not found. Please set them up in the sheet.');
  cache.put('RA_COL_MAP', JSON.stringify(map), 21600);
  return map;
}

function getColumnMap() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('COL_MAP');
  if (cached) return JSON.parse(cached);
  const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const map = {};
  // Layout kolom sama persis di semua sheet tahun MOC Portal (hasil duplikat
  // identik satu sama lain), jadi tidak perlu diikat ke nama sheet tertentu —
  // ambil saja semua named range apa adanya. Ini juga artinya file ini tidak
  // perlu diubah lagi kalau MOC Portal bikin sheet tahun baru (2027, dst).
  ss.getNamedRanges().forEach(nr => {
    try { map[nr.getName().trim()] = nr.getRange().getColumn(); } catch(e) {}
  });
  cache.put('COL_MAP', JSON.stringify(map), 21600);
  return map;
}

function debugUpdateStatus() {
  const testRaId = 'RA/1KtAX2oYFUJiH-wzP9x6ibVkGpsG3ekWr';
  Logger.log('Testing raId: ' + testRaId);
  const result = findMainSheetRowByRaId_(testRaId);
  Logger.log('findMainSheetRowByRaId_ result: ' + JSON.stringify(result));
  if (!result) { Logger.log('ERROR: Row not found!'); return; }
  Logger.log('Row number: ' + result.rowNumber + ' (sheet: ' + result.sheet.getName() + ')');
  const col = getColumnMap();
  Logger.log('Status_Kajian_Risiko column: ' + col.Status_Kajian_Risiko);
  const rowData = result.sheet.getRange(result.rowNumber, 1, 1, result.sheet.getLastColumn()).getValues()[0];
  Logger.log('RA_ID value in row: ' + rowData[col.RA_ID - 1]);
  Logger.log('Status_Kajian_Risiko value: ' + rowData[col.Status_Kajian_Risiko - 1]);
}

function clearAllCaches() {
  const cache = CacheService.getScriptCache();
  cache.remove('COL_MAP');
  cache.remove('RA_COL_MAP');
  cache.remove('RA_CONFIG');
}
