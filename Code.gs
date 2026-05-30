/**
 * ============================================================================
 *  SISTEM MANAJEMEN KEUANGAN KONTER PULSA / DIGITAL PAYMENT
 *  Google Apps Script Generator  -  Code.gs (konfigurasi + orchestrator)
 * ----------------------------------------------------------------------------
 *  Cara pakai singkat:
 *    1. Buka Google Sheets baru  ->  Extensions / Ekstensi  ->  Apps Script.
 *    2. Salin SEMUA file .gs ke project (Code.gs, Structure.gs, dst).
 *    3. Jalankan fungsi  buildAll()  satu kali  (izinkan otorisasi).
 *    4. Reload Sheet  ->  muncul menu "⚙️ Keuangan".
 *
 *  Catatan: semua formula & dropdown aktif s/d baris 5000.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
//  KONFIGURASI GLOBAL
// ---------------------------------------------------------------------------
var CFG = {
  MAX_ROW: 5000,           // baris aktif formula & validasi
  HEADER_ROW: 1,

  // Nama sheet (snake_case, siap jadi tabel/endpoint web app)
  SHEET: {
    PENJUALAN:     'penjualan',
    PEMBELIAN:     'pembelian',
    PRODUK:        'produk',
    TRANSAKSI_KAS: 'transaksi_kas',
    KOREKSI_STOK:  'koreksi_stok',
    UTANG_PIUTANG: 'utang_piutang',
    BUKU_BESAR:    'buku_besar',
    AKUN:          'akun',
    DATA_MASTER:   'data_master',
    DROPDOWN:      '_dropdown',
    KALKULASI:     '_kalkulasi',
    DASHBOARD:     'dashboard',
    LAPORAN:       'laporan',
    AUDIT:         'audit',
    SKENARIO:      'skenario_uji',
    LOG:           'log_perubahan'
  },

  // Palet warna (modern, profesional - bukan kuning polos)
  COLOR: {
    HEADER_BG:   '#102A43',  // biru tua header tabel
    HEADER_TX:   '#FFFFFF',
    BG_APP:      '#F8FAFC',  // abu sangat muda
    CARD:        '#FFFFFF',
    BORDER:      '#E2E8F0',
    ACCENT:      '#1F6FEB',  // biru aksen
    POS:         '#16A34A',  // hijau
    NEG:         '#DC2626',  // merah
    WARN:        '#D97706',  // oranye
    MUTED:       '#64748B',  // abu teks
    INPUT_BG:    '#FFFFFF',  // kolom input (bisa diedit)
    AUTO_BG:     '#F1F5F9',  // kolom formula (diproteksi)
    SNAP_BG:     '#FEF3C7'   // kolom snapshot (diisi otomatis)
  }
};

// Header tiap sheet (urutan = urutan kolom A,B,C,...)
var HEADERS = {
  penjualan: ['Tanggal','ID Transaksi','Jenis Transaksi','Nama Produk','Kode Produk',
    'Kategori Produk','Qty','Harga Modal Snapshot','Harga Jual Snapshot','Total HPP',
    'Total Jual','Laba Kotor','Metode Pembayaran','Akun Uang Masuk','Akun Modal / Sumber HPP',
    'Status Transaksi','Nama Pelanggan','Catatan','User/Admin','Status Closing','Warning'],

  pembelian: ['Tanggal','ID Pembelian','Nama Supplier','Nama Produk','Kode Produk',
    'Kategori Produk','Qty Masuk','Harga Satuan','Total Pembelian','Metode Pembayaran',
    'Akun Pembayaran','Status Pembayaran','Jenis Pembelian','Hutang Supplier','Catatan',
    'User/Admin','Warning'],

  produk: ['Kode Produk','Nama Produk','Kategori','Satuan','Harga Modal Default',
    'Harga Jual Default','Stok Awal','Stok Masuk','Stok Keluar','Koreksi Stok','Stok Akhir',
    'Nilai Stok','Minimal Stok','Status Stok','Warning'],

  transaksi_kas: ['Tanggal','ID','Jenis','Kategori','Akun Keluar','Akun Masuk','Nominal Pokok',
    'Admin Fee','Laba Admin','Metode','Status','Arah Koreksi','Pihak Terkait','Keterangan',
    'User/Admin','Warning'],

  koreksi_stok: ['Tanggal','ID Koreksi','Nama Produk','Kode Produk','Arah','Qty','Alasan',
    'Status','User/Admin','Warning'],

  utang_piutang: ['Tanggal','ID','Arah','Pihak','Nominal','Jatuh Tempo','Sudah Dibayar','Sisa',
    'Akun Bayar','Status','Status Jatuh Tempo','Keterangan','Warning'],

  buku_besar: ['Tanggal','ID Referensi','Sumber Sheet','Jenis Transaksi','Akun','Masuk','Keluar',
    'Saldo Berjalan','Kategori','Status','Keterangan','User/Admin'],

  akun: ['Nama Akun','Jenis Akun','Saldo Awal','Total Masuk','Total Keluar','Saldo Sistem',
    'Saldo Fisik','Selisih','Status Rekonsiliasi','Warning'],

  data_master: ['Status','Final? (Ya/Tidak)','Kategori Produk','Kategori Pengeluaran',
    'Jenis Transaksi Jual','Jenis Transaksi Kas','Jenis Mutasi','Jenis Koreksi Kas',
    'Jenis Koreksi Stok','Metode Pembayaran','Supplier','Pelanggan','User/Admin'],

  _dropdown: ['Produk','Akun','Status','Metode','Supplier','Pelanggan','User','Kategori Produk',
    'Jenis Jual','Jenis Kas','Arah Stok','Arah Utang/Piutang','Jenis Pembelian','Mode Waktu',
    'Arah Koreksi Kas','Kategori Kas','Jenis Akun'],

  _kalkulasi: ['Status Final','Periode Aktif','Label','Nilai Omzet','Nilai Laba','Nilai Beban'],

  laporan: ['Periode','Omzet','HPP','Laba Kotor','Admin Fee','Pengeluaran','Laba Bersih'],

  audit: ['No','Pemeriksaan','Jumlah Temuan','Status','Detail'],

  skenario_uji: ['No','Skenario','Ekspektasi','Hasil Aktual','Verdict'],

  log_perubahan: ['Tanggal','Waktu','User','Sheet','Aksi','Keterangan']
};

// Daftar sheet yang disembunyikan setelah selesai
var HIDDEN_SHEETS = ['_dropdown','_kalkulasi'];

// ---------------------------------------------------------------------------
//  MENU
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Keuangan')
    .addItem('🔄 Refresh Buku Besar & Saldo', 'rebuildLedger')
    .addItem('🧪 Jalankan Skenario Uji', 'runTestScenarios')
    .addSeparator()
    .addItem('🏗️ BUILD ULANG SEMUA (reset)', 'confirmBuildAll')
    .addItem('🔓 Lepas Semua Proteksi', 'removeAllProtections')
    .addToUi();
}

function confirmBuildAll() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('Build Ulang Sistem',
    'Ini akan MENGHAPUS isi & membangun ulang semua sheet. Lanjutkan?',
    ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) buildAll();
}

// ---------------------------------------------------------------------------
//  ORCHESTRATOR UTAMA
// ---------------------------------------------------------------------------
function buildAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Mulai membangun sistem...', 'Build', 5);

  removeAllProtections();        // bersih dulu agar bisa ditulis ulang
  createAllSheets(ss);           // Structure.gs
  seedMasterData(ss);            // Seed.gs  (data_master, akun, produk default)
  setupKalkulasi(ss);            // Seed.gs  (FILTER status final + named range)
  applyFormulas(ss);             // Formulas.gs
  applyValidation(ss);           // Validation.gs
  insertDummyData(ss);           // DummyData.gs
  rebuildLedger();               // Ledger.gs
  buildDashboard(ss);            // Dashboard.gs
  buildLaporan(ss);              // Dashboard.gs
  buildAuditDefinitions(ss);     // Formulas.gs
  seedSkenarioUji(ss);           // DummyData.gs
  applyProtection(ss);           // Protection.gs
  finalize(ss);

  ss.toast('Selesai! Sistem siap dipakai.', 'Build', 8);
}

function finalize(ss) {
  // Sembunyikan sheet sistem
  HIDDEN_SHEETS.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (sh) sh.hideSheet();
  });
  // Aktifkan dashboard sebagai sheet pertama
  var dash = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  if (dash) {
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  }
  logChange(ss, 'SISTEM', 'BUILD', 'Build sistem keuangan selesai');
}

// ---------------------------------------------------------------------------
//  UTILITAS BERSAMA
// ---------------------------------------------------------------------------
function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function colLetter_(n) {           // 1 -> A, 27 -> AA
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function headerIndex_(sheetKey, headerName) {   // 1-based
  var arr = HEADERS[sheetKey];
  for (var i = 0; i < arr.length; i++) if (arr[i] === headerName) return i + 1;
  return -1;
}

function dataRangeA1_(sheetKey, headerName) {   // mis. "D2:D5000"
  var c = colLetter_(headerIndex_(sheetKey, headerName));
  return c + '2:' + c + CFG.MAX_ROW;
}

function logChange(ss, sheet, aksi, ket) {
  try {
    var sh = ss.getSheetByName(CFG.SHEET.LOG);
    if (!sh) return;
    var now = new Date();
    sh.appendRow([Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
      Session.getActiveUser().getEmail() || 'system', sheet, aksi, ket]);
  } catch (e) { /* abaikan */ }
}
