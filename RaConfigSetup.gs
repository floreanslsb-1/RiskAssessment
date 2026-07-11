// ============================================================
// RA CONFIG SETUP — RaConfigSetup.gs
// Jalankan setupRaConfig() SEKALI dari Apps Script editor (Run menu).
// Script ini:
//   - Membuat (atau membuat ulang) sheet "RA_Config"
//   - Mengisi seluruh data yang sekarang hardcode di Index.html/Code.gs
//   - Membuat named range penanda tiap tabel (pola sama seperti
//     MOC_COUNTER_TABLE di ConfigService.gs MOC Portal)
// Aman dijalankan ulang — hanya menghapus & membangun ulang sheet
// RA_Config, tidak menyentuh sheet lain (Kajian_Risiko, Template, dst).
// ============================================================

function setupRaConfig() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('RA_Config');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('RA_Config');

  // FIX: dulu ini hapus SEMUA nama berawalan 'RA_', termasuk named range
  // asli sheet Kajian_Risiko (RA_RA_ID, dst) yang juga pakai awalan sama.
  // Sekarang cuma hapus 8 named range yang memang dibuat script ini.
  const RA_CONFIG_RANGE_NAMES = [
    'RA_DOC_NO', 'RA_DOC_REV', 'RA_DOC_EFFECTIVE_DATE', 'RA_MAIN_APP_URL',
    'RA_KATEGORI_TABLE', 'RA_ROP_LABEL_TABLE', 'RA_RS_LABEL_TABLE', 'RA_ASUMSI_TABLE',
  ];
  ss.getNamedRanges().forEach(function(nr) {
    if (RA_CONFIG_RANGE_NAMES.indexOf(nr.getName()) !== -1) nr.remove();
  });

  let row = 1;

  // ── INFO DOKUMEN ──
  sheet.getRange(row, 1).setValue('[INFO DOKUMEN]').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('No_Dokumen');
  sheet.getRange(row, 2).setValue('SMU/IMS-00/06-013');
  ss.setNamedRange('RA_DOC_NO', sheet.getRange(row, 2));
  row++;
  sheet.getRange(row, 1).setValue('Revisi');
  sheet.getRange(row, 2).setValue('01');
  ss.setNamedRange('RA_DOC_REV', sheet.getRange(row, 2));
  row++;
  sheet.getRange(row, 1).setValue('Tanggal_Berlaku');
  sheet.getRange(row, 2).setValue(new Date());
  sheet.getRange(row, 3).setValue('<- placeholder, set tanggal aslinya lewat Admin Panel nanti');
  ss.setNamedRange('RA_DOC_EFFECTIVE_DATE', sheet.getRange(row, 2));
  row++;
  sheet.getRange(row, 1).setValue('Main_App_URL');
  sheet.getRange(row, 2).setValue(CONFIG.MAIN_APP_URL);
  ss.setNamedRange('RA_MAIN_APP_URL', sheet.getRange(row, 2));
  row += 2;

  // ── KATEGORI RISIKO ──
  sheet.getRange(row, 1).setValue('[KATEGORI RISIKO]').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Kategori');
  ss.setNamedRange('RA_KATEGORI_TABLE', sheet.getRange(row, 1));
  row++;
  const kategoriList = ['Quality', 'Halal', 'OS Proses & Human', 'OH', 'ENV'];
  kategoriList.forEach(function(k) {
    sheet.getRange(row, 1).setValue(k);
    row++;
  });
  row++;

  // ── ROP LABEL (berlaku untuk semua kategori) ──
  sheet.getRange(row, 1).setValue('[ROP LABEL]').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Level');
  sheet.getRange(row, 2).setValue('Label');
  ss.setNamedRange('RA_ROP_LABEL_TABLE', sheet.getRange(row, 1, 1, 2));
  row++;
  const ropLabels = [
    [10, 'Kejadian bisa terjadi di setiap shift'],
    [8,  'Kejadian bisa terjadi 1-2 kali dalam 1 bulan'],
    [6,  'Kejadian bisa terjadi 1-2 kali dalam 1 tahun'],
    [4,  'Kejadian bisa terjadi sekali dalam 2 tahun'],
    [1,  'Kejadian kecil kemungkinan terjadi dalam waktu kurang dari 2-5 tahun'],
  ];
  ropLabels.forEach(function(r) {
    sheet.getRange(row, 1).setValue(r[0]);
    sheet.getRange(row, 2).setValue(r[1]);
    row++;
  });
  row++;

  // ── RS LABEL (per kategori) ──
  sheet.getRange(row, 1).setValue('[RS LABEL]').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Kategori');
  sheet.getRange(row, 2).setValue('Level');
  sheet.getRange(row, 3).setValue('Label');
  ss.setNamedRange('RA_RS_LABEL_TABLE', sheet.getRange(row, 1, 1, 3));
  row++;
  const rsLabels = {
    'Quality': {
      100: 'Kejadian recall yang berisiko berdampak pada kesehatan konsumen',
      40:  'Kejadian recall yang tidak berisiko berdampak pada kesehatan konsumen namun berdampak pada brand',
      21:  'Komplain dari konsumen terkait kualitas atau kuantitas produk',
      8:   'Produksi menghasilkan produk tidak sesuai standar QC (In Process Control) Defect Major',
      2:   'Produksi menghasilkan produk tidak sesuai standar QC (In Process Control) Defect Minor',
    },
    'Halal': {
      100: 'Kejadian recall akibat adanya cemaran bahan haram yang berakibat logo halal dicabut',
      40:  'Kejadian recall akibat adanya cemaran bahan haram yang berdampak pada brand',
      21:  'Kejadian recall akibat adanya cemaran bahan haram yang tidak berdampak pada brand',
      8:   'Menyebabkan cemaran najis berat dan sedang ke produk/fasilitas',
      2:   'Menyebabkan cemaran material tidak halal atau belum di-inquiry ke produk/fasilitas',
    },
    'OS Proses & Human': {
      100: 'Bisa mengakibatkan kematian akibat kecelakaan',
      40:  'Dapat menyebabkan kecelakaan parah yang menyebabkan cacat tetap',
      21:  'Dapat menyebabkan kecelakaan dengan kehilangan jam kerja',
      8:   'Dapat menyebabkan kecelakaan tanpa kehilangan waktu kerja: cedera yang memerlukan perawatan medis',
      2:   'Dapat menyebabkan cedera ringan membutuhkan pertolongan pertama',
    },
    'OH': {
      100: 'Bisa mengakibatkan kematian akibat penyakit',
      40:  'Dapat menyebabkan penyakit parah kronis (pneumonia, kebutaan, keguguran, tuli, kanker, dll)',
      21:  'Dapat menyebabkan penyakit yang membuat pekerja tidak dapat kembali ke pekerjaan biasa',
      8:   'Dapat menyebabkan sakit sementara tanpa kehilangan waktu kerja yang memerlukan perawatan medis',
      2:   'Dapat menyebabkan sakit yang membutuhkan pertolongan pertama',
    },
    'ENV': {
      100: 'Terjadi tumpahan cemaran yang masuk ke badan lingkungan dan dapat mengakibatkan issue eksternal sampai ke ranah hukum',
      40:  'Terjadi tumpahan cemaran yang masuk ke badan lingkungan, sudah menimbulkan issue eksternal namun masih dapat diatasi dengan/tanpa bantuan pihak ketiga',
      21:  'Terjadi tumpahan cemaran yang masih dapat diatasi internal dan belum masuk ke badan lingkungan',
      8:   'Terjadi insiden yang berpotensi dapat menyebabkan tumpahan cemaran',
      2:   'Terdapat kondisi tidak safe (unsafe condition dan action) yang berpotensi dapat menimbulkan insiden tumpahan cemaran',
    },
  };
  const rsLevelOrder = [100, 40, 21, 8, 2];
  kategoriList.forEach(function(kat) {
    rsLevelOrder.forEach(function(lvl) {
      sheet.getRange(row, 1).setValue(kat);
      sheet.getRange(row, 2).setValue(lvl);
      sheet.getRange(row, 3).setValue((rsLabels[kat] && rsLabels[kat][lvl]) || '');
      row++;
    });
  });
  row++;

  // ── ASUMSI / BAHAYA (per kategori) ──
  sheet.getRange(row, 1).setValue('[ASUMSI BAHAYA]').setFontWeight('bold');
  row++;
  sheet.getRange(row, 1).setValue('Kategori');
  sheet.getRange(row, 2).setValue('Asumsi');
  ss.setNamedRange('RA_ASUMSI_TABLE', sheet.getRange(row, 1, 1, 2));
  row++;
  const asumsiByKategori = {
    'Quality': [
      'Machine - Alat tidak terkalibrasi / terverifikasi',
      'Material - Bahan / reagen expired',
      'Material - Bahan awal (BB / BK) tidak sesuai spesifikasi',
      'Material - Bahan baku tidak tersedia',
      'Material - Dokumen tidak lengkap / sesuai',
      'Material - Jumlah dan atau jenis bahan atau produk tidak sesuai',
      'Kapasitas storage / area / ruangan overload',
      'Material - Kerusakan kemasan bahan / produk',
      'Methode - Kesalahan methode / cara kerja',
      'Material - Kesalahan pemesanan material / jasa',
      'Material - Keterlambatan kedatangan',
      'Material - Kontaminasi kotoran / najis / bahan tidak halal',
      'Material - Mixed up (ketercampurbauran)',
      'Man - Salah pemberian identitas / status',
      'Methode - Sampel tidak representatif (mewakili)',
      'Man - Tidak dilakukan record / tidak sesuai waktunya',
      'Machine - Tidak sesuai centerline mesin',
      'Material - Tidak sesuai spesifikasi produk (OOS)',
      'Material - Cost proses / bahan tinggi',
      'Machine - Kerusakan equipment, mesin, instrumen',
      'Material - Perubahan identitas bahan / produk',
      'Methode - Perubahan jadwal',
      'Man - Kesalahan penentuan otorisasi',
    ],
    'Halal': [
      'Material - Dokumen pendukung bahan/material expired',
      'Material - Dokumen pendukung bahan/material tidak lengkap',
      'Material - Bahan baku dan bahan kemas primer belum terdaftar ke bahan halal',
      'Material - Bahan baku terkontaminasi najis',
      'Material - Kemasan primer terkontaminasi najis',
      'Machine - Equipment dan storage terpapar kotoran & najis',
      'Material - Kontaminasi dari personil',
      'Material - Terkontaminasi cairan pelumas',
      'Material - Terkontaminasi bahan pembersih / sanitizer',
      'Material - Campur baur dengan bahan/produk tidak halal',
    ],
    'OS Proses & Human': [
      'A-Anggota tubuh terkena pentalan benda berputar',
      'A-Menginjak benda tajam',
      'A-Tangan tergores material',
      'A-Terjepit Benda Bergerak',
      'A-Terjepit Benda Berputar',
      'A-Terjepit material',
      'A-Tersayat Benda Tajam',
      'B-Terbentur',
      'B-Tertimpa Benda Jatuh',
      'C-Tertabrak Forklift',
      'C-Tertabrak Hand Pallet Manual',
      'C-Tertabrak Kendaraan Eksternal',
      'C-Tertabrak Kendaraan Operasional',
      'C-Tertabrak Pallet Mover',
      'D-Terjatuh pada ketinggian berbeda',
      'D-Terjatuh pada ketinggian sama',
      'D-Terpeleset',
      'D-Tersandung',
      'E-Badan terpapar bahan kimia',
      'E-Badan Terpercik Api',
      'E-Ledakan (Heat)',
      'E-Ledakan (Pencampuran Bahan Kimia)',
      'E-Ledakan (Pressure)',
      'E-Ledakan (Proses)',
      'E-Mata Terpercik api',
      'E-Mata terpercik bahan kimia',
      'E-Tangan terpapar bahan kimia',
      'E-Tangan Terpercik api',
      'E-Terkena Permukaan Panas',
      'E-Terpapar Serpihan Tajam',
      'E-Tersengat Listrik',
      'E-Wajah terpapar bahan kimia',
      'F-Kebakaran',
      'Keadaan Darurat',
    ],
    'OH': [
      'Ergonomi-Mengangkat melebihi kapasitas',
      'Ergonomi-Posisi salah saat mengangkat',
      'Ergonomi-Posisi tubuh Aneh/Canggung',
      'Ergonomi-Repetisi pekerjaan',
      'Fisika-ISBB (Suhu Tubuh)',
      'Fisika-ISBB (Terpapar panas)',
      'Fisika-Pencahayaan Tidak Memadai',
      'Fisika-Suhu Lingkungan Udara Dingin',
      'Fisika-Suhu Lingkungan Udara Panas',
      'Fisika-Terpapar getaran',
      'Fisika-Terpapar Kebisingan',
      'Fisika-Terpapar Radiasi',
      'Fisika-Terpapar Radiasi Elektro Magnetik',
      'Kimia-Terpapar Debu',
      'Kimia-Terpapar kebauan area',
      'Kimia-Terhirup asap',
      'Biologi-Terpapar oleh mikroorganisme',
      'Psikologi-Beban Kerja Berlebih',
      'Fisika-Terpapar Debu',
      'Kimia-Terpapar Uap',
      'Kimia-Terpapar Gas',
    ],
    'ENV': [
      'B3-Pencemaran bahan kimia (B3) cair',
      'B3-Pencemaran bahan kimia (B3) padat',
      'EA-Penggunaan Air Bersih',
      'EF-Penggunaan Bahan Bakar Solar',
      'EG-Penggunaan Gas LPG/Gas PGN',
      'EL-Penggunaan Energi Listrik',
      'LB3-Pencemaran Limbah B3 Cair',
      'LB3-Pencemaran Limbah B3 Padat',
      'PPA-Pencemaran material cair ke badan air',
      'PPA-Pencemaran material cair ke drainase',
      'PPA-Pencemaran material padat ke drainase',
      'PPU-Pencemaran Udara dari Sumber bergerak',
      'PPU-Pencemaran Udara dari Sumber Tidak Bergerak',
      'PPU-Tercemar Kebauan Lingkungan',
      'PPU-Tercemar Kebisingan Lingkungan',
      'SP-Menimbulkan Sampah recycle (plastik, kertas)',
      'SP-Menimbulkan Sampah Residu (sampah yang tidak bisa di daur ulang)',
    ],
  };
  kategoriList.forEach(function(kat) {
    (asumsiByKategori[kat] || []).forEach(function(a) {
      sheet.getRange(row, 1).setValue(kat);
      sheet.getRange(row, 2).setValue(a);
      row++;
    });
  });

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 420);

  SpreadsheetApp.flush();
  Logger.log('RA_Config setup selesai. Total baris terpakai: ' + row);
}
