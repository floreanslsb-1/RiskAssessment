// ============================================================
// RA CONFIG ADMIN — RaConfigAdmin.gs
// Backend untuk Admin Panel. Setiap fungsi save* membaca config
// yang ada (getRaConfig), menggabungkan perubahan, lalu menulis
// ULANG seluruh sheet RA_Config dari nol (writeRaConfigToSheet_).
// Ini menghindari bug geser-baris kalau jumlah kategori/asumsi
// berubah ukuran.
// ============================================================

// ── AKSES ──────────────────────────────────────────────────
function isRaAdmin_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const nr = ss.getRangeByName('RA_ADMIN_EMAILS');
    if (!nr) return false;
    const list = String(nr.getValue() || '')
      .split(',').map(function(e) { return e.trim().toLowerCase(); }).filter(Boolean);
    const me = Session.getActiveUser().getEmail().toLowerCase();
    return list.indexOf(me) !== -1;
  } catch (e) { return false; }
}

// Dipanggil SEKALI secara manual untuk mengisi admin pertama.
// Ganti 'email-kamu@wingscorp.com' lalu jalankan sekali dari editor.
function bootstrapRaAdmin() {
  const email = 'floreansalsabila.irdana@wingscorp.com'; // <-- GANTI INI
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('RA_Config');
  if (!sheet) throw new Error('Sheet RA_Config belum ada. Jalankan setupRaConfig() dulu.');
  let nr = ss.getRangeByName('RA_ADMIN_EMAILS');
  if (!nr) {
    const row = sheet.getLastRow() + 2;
    sheet.getRange(row, 1).setValue('Admin_Emails');
    nr = sheet.getRange(row, 2);
    ss.setNamedRange('RA_ADMIN_EMAILS', nr);
  }
  nr.setValue(email);
  Logger.log('Admin RA app: ' + email);
}

function getRaAdminData() {
  return {
    isAdmin: isRaAdmin_(),
    config: getRaConfig(),
  };
}

// ── PENULIS ULANG SHEET (dipakai semua fungsi save*) ────────
function writeRaConfigToSheet_(config) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const oldSheet = ss.getSheetByName('RA_Config');

  // Simpan RA_ADMIN_EMAILS dulu sebelum sheet dihapus, biar ikut dipindah
  let adminEmailsValue = '';
  try { adminEmailsValue = ss.getRangeByName('RA_ADMIN_EMAILS').getValue(); } catch (e) {}

  if (oldSheet) ss.deleteSheet(oldSheet);
  const sheet = ss.insertSheet('RA_Config');

  ['RA_DOC_NO','RA_DOC_REV','RA_DOC_EFFECTIVE_DATE','RA_MAIN_APP_URL',
   'RA_KATEGORI_TABLE','RA_ROP_LABEL_TABLE','RA_RS_LABEL_TABLE','RA_ASUMSI_TABLE',
   'RA_ADMIN_EMAILS'].forEach(function(name) {
    try { ss.removeNamedRange(name); } catch (e) {}
  });

  let row = 1;

  sheet.getRange(row, 1).setValue('[INFO DOKUMEN]').setFontWeight('bold'); row++;
  sheet.getRange(row, 1).setValue('No_Dokumen');
  sheet.getRange(row, 2).setValue(config.docNo);
  ss.setNamedRange('RA_DOC_NO', sheet.getRange(row, 2)); row++;
  sheet.getRange(row, 1).setValue('Revisi');
  sheet.getRange(row, 2).setValue(config.docRev);
  ss.setNamedRange('RA_DOC_REV', sheet.getRange(row, 2)); row++;
  sheet.getRange(row, 1).setValue('Tanggal_Berlaku');
  sheet.getRange(row, 2).setValue(config.effectiveDate);
  ss.setNamedRange('RA_DOC_EFFECTIVE_DATE', sheet.getRange(row, 2)); row++;
  sheet.getRange(row, 1).setValue('Main_App_URL');
  sheet.getRange(row, 2).setValue(config.mainAppUrl);
  ss.setNamedRange('RA_MAIN_APP_URL', sheet.getRange(row, 2)); row++;
  sheet.getRange(row, 1).setValue('Admin_Emails');
  sheet.getRange(row, 2).setValue(adminEmailsValue);
  ss.setNamedRange('RA_ADMIN_EMAILS', sheet.getRange(row, 2)); row += 2;

  sheet.getRange(row, 1).setValue('[KATEGORI RISIKO]').setFontWeight('bold'); row++;
  sheet.getRange(row, 1).setValue('Kategori');
  ss.setNamedRange('RA_KATEGORI_TABLE', sheet.getRange(row, 1)); row++;
  config.kategoriList.forEach(function(k) {
    sheet.getRange(row, 1).setValue(k); row++;
  });
  row++;

  sheet.getRange(row, 1).setValue('[ROP LABEL]').setFontWeight('bold'); row++;
  sheet.getRange(row, 1).setValue('Level');
  sheet.getRange(row, 2).setValue('Label');
  ss.setNamedRange('RA_ROP_LABEL_TABLE', sheet.getRange(row, 1, 1, 2)); row++;
  [10, 8, 6, 4, 1].forEach(function(lvl) {
    sheet.getRange(row, 1).setValue(lvl);
    sheet.getRange(row, 2).setValue((config.ropLabels && config.ropLabels[lvl]) || '');
    row++;
  });
  row++;

  sheet.getRange(row, 1).setValue('[RS LABEL]').setFontWeight('bold'); row++;
  sheet.getRange(row, 1).setValue('Kategori');
  sheet.getRange(row, 2).setValue('Level');
  sheet.getRange(row, 3).setValue('Label');
  ss.setNamedRange('RA_RS_LABEL_TABLE', sheet.getRange(row, 1, 1, 3)); row++;
  config.kategoriList.forEach(function(kat) {
    [100, 40, 21, 8, 2].forEach(function(lvl) {
      sheet.getRange(row, 1).setValue(kat);
      sheet.getRange(row, 2).setValue(lvl);
      sheet.getRange(row, 3).setValue(
        (config.rsLabelsByKategori[kat] && config.rsLabelsByKategori[kat][lvl]) || ''
      );
      row++;
    });
  });
  row++;

  sheet.getRange(row, 1).setValue('[ASUMSI BAHAYA]').setFontWeight('bold'); row++;
  sheet.getRange(row, 1).setValue('Kategori');
  sheet.getRange(row, 2).setValue('Asumsi');
  ss.setNamedRange('RA_ASUMSI_TABLE', sheet.getRange(row, 1, 1, 2)); row++;
  config.kategoriList.forEach(function(kat) {
    (config.asumsiByKategori[kat] || []).forEach(function(a) {
      sheet.getRange(row, 1).setValue(kat);
      sheet.getRange(row, 2).setValue(a);
      row++;
    });
  });

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 420);
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove('RA_CONFIG');
}

// ── FUNGSI SIMPAN (dipanggil dari Admin Panel) ──────────────
function saveRaInfoDokumen(data) {
  if (!isRaAdmin_()) return { success: false, error: 'Akses ditolak.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getRaConfig();
    config.docNo = data.docNo;
    config.docRev = data.docRev;
    config.effectiveDate = data.effectiveDate ? new Date(data.effectiveDate) : config.effectiveDate;
    config.mainAppUrl = data.mainAppUrl;
    writeRaConfigToSheet_(config);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function saveRaKategoriList(newList) {
  if (!isRaAdmin_()) return { success: false, error: 'Akses ditolak.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getRaConfig();
    config.kategoriList = newList;
    writeRaConfigToSheet_(config);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function saveRaLabels(ropLabels, rsLabelsByKategori) {
  if (!isRaAdmin_()) return { success: false, error: 'Akses ditolak.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getRaConfig();
    config.ropLabels = ropLabels;
    config.rsLabelsByKategori = rsLabelsByKategori;
    writeRaConfigToSheet_(config);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function saveRaAsumsiForKategori(kategori, items) {
  if (!isRaAdmin_()) return { success: false, error: 'Akses ditolak.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getRaConfig();
    if (!config.asumsiByKategori) config.asumsiByKategori = {};
    config.asumsiByKategori[kategori] = items;
    writeRaConfigToSheet_(config);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function getRaAdminEmailsList_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const nr = ss.getRangeByName('RA_ADMIN_EMAILS');
    if (!nr) return [];
    return String(nr.getValue() || '').split(',').map(function(e) { return e.trim(); }).filter(Boolean);
  } catch (e) { return []; }
}

function saveRaAdminEmails(emailsList) {
  if (!isRaAdmin_()) return { success: false, error: 'Akses ditolak.' };
  const cleaned = (emailsList || [])
    .map(function(e) { return String(e).trim().toLowerCase(); })
    .filter(Boolean);
  if (cleaned.length === 0) {
    return { success: false, error: 'Minimal harus ada 1 admin.' };
  }
  const me = Session.getActiveUser().getEmail().toLowerCase();
  if (cleaned.indexOf(me) === -1) {
    return { success: false, error: 'Kamu tidak bisa menghapus emailmu sendiri dari daftar admin.' };
  }
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const nr = ss.getRangeByName('RA_ADMIN_EMAILS');
  if (!nr) return { success: false, error: 'Named range RA_ADMIN_EMAILS belum ada.' };
  nr.setValue(cleaned.join(','));
  return { success: true };
}
