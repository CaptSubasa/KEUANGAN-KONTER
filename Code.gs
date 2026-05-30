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
  var ui = SpreadsheetApp.getUi();

  var subTambah = ui.createMenu('➕ Tambah Cepat')
    .addItem('Tambah Produk Baru', 'tambahProduk')
    .addItem('Tambah Akun Baru', 'tambahAkun');

  var subReset = ui.createMenu('🧹 Reset / Hapus Data')
    .addItem('Hapus Semua Data Transaksi', 'confirmHapusDataContoh')
    .addItem('Reset Daftar Produk', 'confirmResetProduk')
    .addItem('Reset Daftar Akun', 'confirmResetAkun')
    .addItem('Reset SEMUA (transaksi + produk + akun)', 'confirmResetTotal');

  ui.createMenu('⚙️ Keuangan')
    .addItem('🔄 Refresh Buku Besar & Saldo', 'rebuildLedger')
    .addItem('🔒 Bekukan Harga Snapshot Penjualan', 'freezeAllSnapshots')
    .addItem('🧪 Jalankan Skenario Uji', 'runTestScenarios')
    .addSeparator()
    .addSubMenu(subTambah)
    .addSubMenu(subReset)
    .addItem('💰 Atur Saldo Awal Akun', 'aturSaldoAwal')
    .addSeparator()
    .addItem('🏗️ BUILD ULANG SEMUA (reset total)', 'confirmBuildAll')
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



// ===========================================================================
//  HAPUS DATA CONTOH  (siap pakai data asli)
// ===========================================================================
function confirmHapusDataContoh() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('Hapus Semua Data Contoh',
    'Ini akan MENGHAPUS semua data contoh di:\n\n' +
    '• penjualan\n• pembelian\n• transaksi_kas\n• koreksi_stok\n• utang_piutang\n• buku_besar\n• log_perubahan\n\n' +
    'Master tetap aman: produk, akun, data_master.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  hapusDataContoh();
}

function hapusDataContoh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Membersihkan data contoh...', 'Hapus', 4);

  // Sheet transaksi -> kosongkan baris 2 ke bawah, kolom A sampai akhir header
  var sheets = [
    CFG.SHEET.PENJUALAN, CFG.SHEET.PEMBELIAN, CFG.SHEET.TRANSAKSI_KAS,
    CFG.SHEET.KOREKSI_STOK, CFG.SHEET.UTANG_PIUTANG, CFG.SHEET.LOG
  ];
  sheets.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var nCol = HEADERS[name] ? HEADERS[name].length : sh.getLastColumn();
    sh.getRange(2, 1, CFG.MAX_ROW - 1, nCol).clearContent();
  });

  // Buku besar dibangun ulang otomatis
  rebuildLedger();
  SpreadsheetApp.flush();

  // Hapus saldo fisik (G) di akun supaya bersih, biarkan saldo awal
  var ak = ss.getSheetByName(CFG.SHEET.AKUN);
  if (ak) ak.getRange('G2:G' + CFG.MAX_ROW).clearContent();

  logChange(ss, 'SISTEM', 'BERSIH', 'Semua data contoh dihapus, siap pakai data asli');
  SpreadsheetApp.getUi().alert('✅ Selesai',
    'Semua data contoh telah dihapus.\n\n' +
    'LANGKAH BERIKUTNYA:\n' +
    '1. Buka sheet "akun" → atur Saldo Awal sesuai uang Anda yang sebenarnya.\n' +
    '2. Mulai input transaksi harian di sheet "penjualan", "pembelian", "transaksi_kas".\n' +
    '3. Klik ⚙️ Keuangan → 🔄 Refresh setiap selesai input.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===========================================================================
//  ATUR ULANG SALDO AWAL  (template kosong supaya user isi manual)
// ===========================================================================
function aturSaldoAwal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ak = ss.getSheetByName(CFG.SHEET.AKUN);
  if (!ak) return;
  ss.setActiveSheet(ak);
  ak.getRange('C2').activate();
  SpreadsheetApp.getUi().alert('💰 Atur Saldo Awal',
    'Saya pindahkan Anda ke sheet "akun".\n\n' +
    'Isi kolom "Saldo Awal" (kolom C) sesuai uang Anda saat ini di tiap akun:\n' +
    '- Kas Utama (uang tunai)\n' +
    '- Bank BCA (saldo rekening)\n' +
    '- QRIS, ShopeePay, GoPay, DANA, OVO (saldo e-wallet)\n' +
    '- Saldo Supplier (deposit di supplier pulsa)\n\n' +
    'Setelah selesai, klik ⚙️ Keuangan → 🔄 Refresh.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}



// ===========================================================================
//  RESET PRODUK
// ===========================================================================
function confirmResetProduk() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('🧹 Reset Daftar Produk',
    'HAPUS semua produk di sheet "produk"?\n\n' +
    '⚠️ Setelah reset, dropdown produk di penjualan/pembelian akan kosong sampai Anda input produk baru.\n\n' +
    'Lanjutkan?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  resetProduk();
  ui.alert('✅ Selesai',
    'Daftar produk dikosongkan.\n\nIsi produk baru via:\n' +
    '• Menu ⚙️ Keuangan → ➕ Tambah Cepat → Tambah Produk Baru, atau\n' +
    '• Langsung ketik di sheet "produk" (Kode, Nama, Kategori, Modal, Jual, Stok Awal).',
    ui.ButtonSet.OK);
}

function resetProduk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);
  if (!pr) return;
  pr.getRange(2, 1, CFG.MAX_ROW - 1, HEADERS.produk.length).clearContent();
  logChange(ss, 'produk', 'RESET', 'Daftar produk dikosongkan');
  ss.setActiveSheet(pr); pr.getRange('A2').activate();
}

// ===========================================================================
//  RESET AKUN
// ===========================================================================
function confirmResetAkun() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('🧹 Reset Daftar Akun',
    'HAPUS semua akun di sheet "akun"?\n\n' +
    '⚠️ Setelah reset, dropdown akun (Kas Utama, Bank, dll) akan kosong sampai Anda input akun baru.\n\n' +
    'Lanjutkan?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  resetAkun();
  ui.alert('✅ Selesai',
    'Daftar akun dikosongkan.\n\nIsi akun baru via:\n' +
    '• Menu ⚙️ Keuangan → ➕ Tambah Cepat → Tambah Akun Baru, atau\n' +
    '• Langsung ketik di sheet "akun" (Nama, Jenis, Saldo Awal).',
    ui.ButtonSet.OK);
}

function resetAkun() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ak = ss.getSheetByName(CFG.SHEET.AKUN);
  if (!ak) return;
  ak.getRange(2, 1, CFG.MAX_ROW - 1, HEADERS.akun.length).clearContent();
  logChange(ss, 'akun', 'RESET', 'Daftar akun dikosongkan');
  ss.setActiveSheet(ak); ak.getRange('A2').activate();
}

// ===========================================================================
//  RESET SEMUA (transaksi + produk + akun)
// ===========================================================================
function confirmResetTotal() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('⚠️ Reset SEMUA',
    'INI AKAN MENGHAPUS:\n' +
    '• Semua transaksi (penjualan, pembelian, transaksi_kas, dll)\n' +
    '• Semua produk\n' +
    '• Semua akun\n' +
    '• Buku besar & log\n\n' +
    'Yang TIDAK dihapus: data_master (status, kategori, supplier, dll).\n\n' +
    'Yakin lanjut?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  hapusDataContoh();
  resetProduk();
  resetAkun();
  SpreadsheetApp.getActiveSpreadsheet().toast('Semua data dihapus. Mulai dari nol.', '✅ Reset Total', 6);
  ui.alert('✅ Reset Total Selesai',
    'Sistem kembali kosong.\n\nLANGKAH MULAI:\n' +
    '1. ➕ Tambah Akun Baru (Kas, Bank, QRIS, dll) + saldo awal\n' +
    '2. ➕ Tambah Produk Baru (pulsa, paket data, dll)\n' +
    '3. Mulai input transaksi harian\n' +
    '4. 🔄 Refresh setiap selesai input',
    ui.ButtonSet.OK);
}

// ===========================================================================
//  TAMBAH PRODUK CEPAT (form via prompt)
// ===========================================================================
function tambahProduk() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);

  var p1 = ui.prompt('➕ Tambah Produk - 1/6', 'Kode Produk (mis. PUL5):', ui.ButtonSet.OK_CANCEL);
  if (p1.getSelectedButton() !== ui.Button.OK || !p1.getResponseText().trim()) return;
  var p2 = ui.prompt('➕ Tambah Produk - 2/6', 'Nama Produk (mis. Pulsa 5K):', ui.ButtonSet.OK_CANCEL);
  if (p2.getSelectedButton() !== ui.Button.OK || !p2.getResponseText().trim()) return;
  var p3 = ui.prompt('➕ Tambah Produk - 3/6',
    'Kategori. Pilih salah satu:\n• Pulsa\n• Paket Data\n• Voucher\n• Token Listrik\n• E-Wallet\n• Layanan Digital',
    ui.ButtonSet.OK_CANCEL);
  if (p3.getSelectedButton() !== ui.Button.OK) return;
  var p4 = ui.prompt('➕ Tambah Produk - 4/6', 'Harga Modal Default (angka, mis. 5500):', ui.ButtonSet.OK_CANCEL);
  if (p4.getSelectedButton() !== ui.Button.OK) return;
  var p5 = ui.prompt('➕ Tambah Produk - 5/6', 'Harga Jual Default (angka, mis. 6500):', ui.ButtonSet.OK_CANCEL);
  if (p5.getSelectedButton() !== ui.Button.OK) return;
  var p6 = ui.prompt('➕ Tambah Produk - 6/6', 'Stok Awal (angka, mis. 100. Untuk pulsa/PPOB tanpa stok fisik isi 0):', ui.ButtonSet.OK_CANCEL);
  if (p6.getSelectedButton() !== ui.Button.OK) return;

  var row = nextEmptyRow_(pr, 1);
  pr.getRange(row, 1, 1, 7).setValues([[
    p1.getResponseText().trim().toUpperCase(),
    p2.getResponseText().trim(),
    (p3.getResponseText().trim() || 'Pulsa'),
    'transaksi',
    Number(p4.getResponseText()) || 0,
    Number(p5.getResponseText()) || 0,
    Number(p6.getResponseText()) || 0
  ]]);
  pr.getRange(row, 13).setValue(5); // minimal stok default 5
  ss.toast('Produk "' + p2.getResponseText().trim() + '" ditambahkan ✅', 'Sukses', 5);
  logChange(ss, 'produk', 'TAMBAH', p2.getResponseText().trim());
}

// ===========================================================================
//  TAMBAH AKUN CEPAT
// ===========================================================================
function tambahAkun() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ak = ss.getSheetByName(CFG.SHEET.AKUN);

  var p1 = ui.prompt('➕ Tambah Akun - 1/3', 'Nama Akun (mis. Bank Mandiri):', ui.ButtonSet.OK_CANCEL);
  if (p1.getSelectedButton() !== ui.Button.OK || !p1.getResponseText().trim()) return;
  var p2 = ui.prompt('➕ Tambah Akun - 2/3',
    'Jenis Akun. Pilih salah satu:\n• Kas\n• Bank\n• E-Wallet\n• QRIS\n• Supplier',
    ui.ButtonSet.OK_CANCEL);
  if (p2.getSelectedButton() !== ui.Button.OK) return;
  var p3 = ui.prompt('➕ Tambah Akun - 3/3', 'Saldo Awal (angka, mis. 1000000):', ui.ButtonSet.OK_CANCEL);
  if (p3.getSelectedButton() !== ui.Button.OK) return;

  var row = nextEmptyRow_(ak, 1);
  ak.getRange(row, 1, 1, 3).setValues([[
    p1.getResponseText().trim(),
    (p2.getResponseText().trim() || 'Kas'),
    Number(p3.getResponseText()) || 0
  ]]);
  ss.toast('Akun "' + p1.getResponseText().trim() + '" ditambahkan ✅', 'Sukses', 5);
  logChange(ss, 'akun', 'TAMBAH', p1.getResponseText().trim());
  rebuildLedger();
}

/** Cari baris kosong pertama berdasarkan kolom kunci */
function nextEmptyRow_(sh, keyCol) {
  var last = sh.getLastRow();
  if (last < 1) return 2;
  var vals = sh.getRange(2, keyCol, Math.max(last, CFG.MAX_ROW) - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === '' || vals[i][0] === null) return i + 2;
  }
  return Math.min(last + 1, CFG.MAX_ROW);
}
