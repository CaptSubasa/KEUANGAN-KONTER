var CFG = {
  MAX_ROW: 5000,           // baris aktif formula & validasi
  HEADER_ROW: 1,
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
var HIDDEN_SHEETS = ['_dropdown','_kalkulasi'];
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
  HIDDEN_SHEETS.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (sh) sh.hideSheet();
  });
  var dash = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  if (dash) {
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  }
  logChange(ss, 'SISTEM', 'BUILD', 'Build sistem keuangan selesai');
}
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
  rebuildLedger();
  SpreadsheetApp.flush();
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
function nextEmptyRow_(sh, keyCol) {
  var last = sh.getLastRow();
  if (last < 1) return 2;
  var vals = sh.getRange(2, keyCol, Math.max(last, CFG.MAX_ROW) - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === '' || vals[i][0] === null) return i + 2;
  }
  return Math.min(last + 1, CFG.MAX_ROW);
}
function createAllSheets(ss) {
  var order = [
    CFG.SHEET.DASHBOARD, CFG.SHEET.PENJUALAN, CFG.SHEET.PEMBELIAN, CFG.SHEET.PRODUK,
    CFG.SHEET.TRANSAKSI_KAS, CFG.SHEET.KOREKSI_STOK, CFG.SHEET.UTANG_PIUTANG,
    CFG.SHEET.BUKU_BESAR, CFG.SHEET.AKUN, CFG.SHEET.LAPORAN, CFG.SHEET.AUDIT,
    CFG.SHEET.SKENARIO, CFG.SHEET.DATA_MASTER, CFG.SHEET.LOG,
    CFG.SHEET.DROPDOWN, CFG.SHEET.KALKULASI
  ];
  order.forEach(function(name, i) {
    var sh = getOrCreateSheet_(ss, name);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (def && order.indexOf(def.getName()) === -1) ss.deleteSheet(def);
  Object.keys(HEADERS).forEach(function(key) {
    if (key === '_kalkulasi') return; // _kalkulasi ditata di Seed.gs
    writeHeader_(ss, key, HEADERS[key]);
  });
  ['penjualan','pembelian','produk','transaksi_kas','koreksi_stok','utang_piutang',
   'buku_besar','akun','data_master','laporan','audit','skenario_uji','log_perubahan','_dropdown']
    .forEach(function(key){ styleDataSheet_(ss, key); });
}
function writeHeader_(ss, sheetKey, headers) {
  var sh = ss.getSheetByName(sheetKey);
  if (!sh) return;
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  var hr = sh.getRange(1, 1, 1, headers.length);
  hr.setBackground(CFG.COLOR.HEADER_BG)
    .setFontColor(CFG.COLOR.HEADER_TX)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);
}
function styleDataSheet_(ss, sheetKey) {
  var sh = ss.getSheetByName(sheetKey);
  if (!sh) return;
  var nCol = HEADERS[sheetKey] ? HEADERS[sheetKey].length : sh.getLastColumn();
  if (!nCol) return;
  for (var c = 1; c <= nCol; c++) {
    sh.setColumnWidth(c, 130);
  }
  var maxCols = sh.getMaxColumns();
  if (maxCols > nCol + 3) sh.deleteColumns(nCol + 4, maxCols - (nCol + 3));
  var maxRows = sh.getMaxRows();
  if (maxRows < CFG.MAX_ROW) sh.insertRowsAfter(maxRows, CFG.MAX_ROW - maxRows);
  sh.setHiddenGridlines ? null : null;
}
function fmtRupiah_(sh, a1) {
  sh.getRange(a1).setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0');
}
function fmtTanggal_(sh, a1) {
  sh.getRange(a1).setNumberFormat('yyyy-mm-dd');
}
function fmtAngka_(sh, a1) {
  sh.getRange(a1).setNumberFormat('#,##0');
}
function seedMasterData(ss) {
  var dm = ss.getSheetByName(CFG.SHEET.DATA_MASTER);
  var status = [
    ['Pending','Tidak'], ['Draft','Tidak'], ['Proses','Tidak'],
    ['Lunas','Ya'], ['Selesai','Ya'], ['Sukses','Ya'], ['Approved','Ya'], ['Disetujui','Ya'],
    ['Batal','Tidak'], ['Gagal','Tidak']
  ];
  dm.getRange(2, 1, status.length, 2).setValues(status);
  setCol_(dm, 3, ['Pulsa','Paket Data','Voucher','Token Listrik','E-Wallet','Layanan Digital']);
  setCol_(dm, 4, ['Listrik','Sewa Tempat','Gaji','Internet','ATK','Transport','Konsumsi','Lain-lain']);
  setCol_(dm, 5, ['Pulsa','Paket Data','Voucher Game','Token Listrik','Top Up E-Wallet','PPOB','Lainnya']);
  setCol_(dm, 6, ['pengeluaran','modal','mutasi','tarik_tunai','transfer_uang','koreksi_kas']);
  setCol_(dm, 7, ['Setor ke Bank','Tarik dari Bank','Isi Saldo E-Wallet','Tarik Saldo E-Wallet','Topup QRIS']);
  setCol_(dm, 8, ['Selisih Lebih','Selisih Kurang','Pembulatan','Koreksi Salah Input']);
  setCol_(dm, 9, ['Barang Rusak','Barang Hilang','Salah Hitung','Stok Opname']);
  setCol_(dm, 10, ['Tunai','Transfer Bank','QRIS','E-Wallet','Kredit']);
  setCol_(dm, 11, ['Supplier Pulsa A','Distributor Data B','PPOB Center C','Agen Token D']);
  setCol_(dm, 12, ['Umum','Budi','Siti','Andi','Rina','Toko Maju','Warung Bu Tini']);
  setCol_(dm, 13, ['Admin','Owner','Kasir 1','Kasir 2']);
  var ak = ss.getSheetByName(CFG.SHEET.AKUN);
  var akun = [
    ['Kas Utama','Kas',500000],
    ['Bank BCA','Bank',2000000],
    ['QRIS','QRIS',0],
    ['ShopeePay','E-Wallet',300000],
    ['GoPay','E-Wallet',250000],
    ['DANA','E-Wallet',150000],
    ['OVO','E-Wallet',100000],
    ['Saldo Supplier','Supplier',1000000]
  ];
  ak.getRange(2, 1, akun.length, 3).setValues(akun);
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);
  var produk = [
    ['PUL5','Pulsa 5K','Pulsa','transaksi',5500,6500,100],
    ['PUL10','Pulsa 10K','Pulsa','transaksi',10500,12000,100],
    ['PUL25','Pulsa 25K','Pulsa','transaksi',24800,26500,60],
    ['DATA1','Paket Data 1GB','Paket Data','transaksi',9000,12000,80],
    ['DATA5','Paket Data 5GB','Paket Data','transaksi',25000,32000,50],
    ['TKN20','Token Listrik 20K','Token Listrik','transaksi',20500,22000,40],
    ['TKN50','Token Listrik 50K','Token Listrik','transaksi',50500,52500,40],
    ['VCRG10','Voucher Game 10K','Voucher','pcs',9500,11000,30],
    ['EWAL50','Top Up E-Wallet 50K','E-Wallet','transaksi',50000,51500,100]
  ];
  pr.getRange(2, 1, produk.length, 7).setValues(produk);
  var minStok = produk.map(function(){ return [10]; });
  pr.getRange(2, 13, minStok.length, 1).setValues(minStok);
}
function setCol_(sh, col, arr) {
  var vals = arr.map(function(v){ return [v]; });
  sh.getRange(2, col, vals.length, 1).setValues(vals);
}
function setupKalkulasi(ss) {
  var k = ss.getSheetByName(CFG.SHEET.KALKULASI);
  k.clear();
  k.getRange('A1').setValue('Status Final').setFontWeight('bold');
  k.getRange('A2').setFormula(
    '=IFERROR(FILTER(' + CFG.SHEET.DATA_MASTER + '!A2:A' + CFG.MAX_ROW + ', ' +
    CFG.SHEET.DATA_MASTER + '!B2:B' + CFG.MAX_ROW + '="Ya"),"")'
  );
  var rng = k.getRange('A2:A60');
  ss.setNamedRange('FINAL_STATUS', rng);
  k.getRange('C1').setValue('(area helper dashboard - lihat Dashboard.gs)')
    .setFontColor(CFG.COLOR.MUTED);
}
var R = function (a1) { return a1; };           // alias keterbacaan
var MR = CFG.MAX_ROW;
function setAF_(sh, a1, formula) { sh.getRange(a1).setFormula(formula); }
function applyFormulas(ss) {
  formulasPenjualan_(ss);
  formulasPembelian_(ss);
  formulasProduk_(ss);
  formulasTransaksiKas_(ss);
  formulasKoreksiStok_(ss);
  formulasUtangPiutang_(ss);
  formulasAkun_(ss);
  applyNumberFormats_(ss);
}
function formulasPenjualan_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.PENJUALAN);
  var P = CFG.SHEET.PRODUK;
  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","TRX-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'E2', '=ARRAYFORMULA(IF(D2:D'+MR+'="","",IFERROR(INDEX('+P+'!$A$2:$A$'+MR+',MATCH(D2:D'+MR+','+P+'!$B$2:$B$'+MR+',0)),"❓ tdk ada")))');
  setAF_(sh, 'F2', '=ARRAYFORMULA(IF(D2:D'+MR+'="","",IFERROR(INDEX('+P+'!$C$2:$C$'+MR+',MATCH(D2:D'+MR+','+P+'!$B$2:$B$'+MR+',0)),"")))');
  setAF_(sh, 'J2', '=ARRAYFORMULA(IF(G2:G'+MR+'="","",N(G2:G'+MR+')*N(H2:H'+MR+')))');
  setAF_(sh, 'K2', '=ARRAYFORMULA(IF(G2:G'+MR+'="","",N(G2:G'+MR+')*N(I2:I'+MR+')))');
  setAF_(sh, 'L2', '=ARRAYFORMULA(IF(G2:G'+MR+'="","",K2:K'+MR+'-J2:J'+MR+'))');
  setAF_(sh, 'T2', '=ARRAYFORMULA(IF(P2:P'+MR+'="","",IF(COUNTIF(FINAL_STATUS,P2:P'+MR+')>0,"Final","Non-Final")))');
  setAF_(sh, 'U2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(D2:D'+MR+'="","⚠️ Produk kosong",'+
    'IF(N(G2:G'+MR+')<=0,"⚠️ Qty kosong/0",'+
    'IF((H2:H'+MR+'=0)+(I2:I'+MR+'=0)>0,"⚠️ Harga belum tersnapshot",'+
    'IF((T2:T'+MR+'="Final")*(N2:N'+MR+'=""),"⚠️ Akun Uang Masuk kosong",'+
    'IF((T2:T'+MR+'="Final")*(O2:O'+MR+'=""),"⚠️ Akun Modal kosong","")))))))');
  sh.getRange('V1').setValue('eff_keluar (sys)');
  setAF_(sh, 'V2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(T2:T'+MR+'="Final",N(G2:G'+MR+'),0)))');
  sh.getRange('W1').setValue('eff_laba (sys)');
  setAF_(sh, 'W2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(T2:T'+MR+'="Final",L2:L'+MR+',0)))');
  sh.hideColumns(22, 2);
}
function formulasPembelian_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.PEMBELIAN);
  var P = CFG.SHEET.PRODUK;
  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","BLI-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'E2', '=ARRAYFORMULA(IF(D2:D'+MR+'="","",IFERROR(INDEX('+P+'!$A$2:$A$'+MR+',MATCH(D2:D'+MR+','+P+'!$B$2:$B$'+MR+',0)),"❓ tdk ada")))');
  setAF_(sh, 'F2', '=ARRAYFORMULA(IF(D2:D'+MR+'="","",IFERROR(INDEX('+P+'!$C$2:$C$'+MR+',MATCH(D2:D'+MR+','+P+'!$B$2:$B$'+MR+',0)),"")))');
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(G2:G'+MR+'="","",N(G2:G'+MR+')*N(H2:H'+MR+')))');
  setAF_(sh, 'N2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(M2:M'+MR+'="Kredit",I2:I'+MR+',0)))');
  setAF_(sh, 'Q2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(D2:D'+MR+'="","⚠️ Produk kosong",'+
    'IF(N(G2:G'+MR+')<=0,"⚠️ Qty kosong/0",'+
    'IF((M2:M'+MR+'<>"Kredit")*(K2:K'+MR+'=""),"⚠️ Akun pembayaran kosong","")))))');
  sh.getRange('R1').setValue('eff_masuk (sys)');
  setAF_(sh, 'R2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(COUNTIF(FINAL_STATUS,L2:L'+MR+')>0,N(G2:G'+MR+'),0)))');
  sh.hideColumns(18);
}
function formulasProduk_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.PRODUK);
  var JUAL = CFG.SHEET.PENJUALAN, BELI = CFG.SHEET.PEMBELIAN, KOR = CFG.SHEET.KOREKSI_STOK;
  setAF_(sh, 'H2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+BELI+'!$E$2:$E$'+MR+',A2:A'+MR+','+BELI+'!$R$2:$R$'+MR+')))');
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+JUAL+'!$E$2:$E$'+MR+',A2:A'+MR+','+JUAL+'!$V$2:$V$'+MR+')))');
  setAF_(sh, 'J2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+KOR+'!$D$2:$D$'+MR+',A2:A'+MR+','+KOR+'!$K$2:$K$'+MR+')))');
  setAF_(sh, 'K2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",N(G2:G'+MR+')+H2:H'+MR+'-I2:I'+MR+'+J2:J'+MR+'))');
  setAF_(sh, 'L2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",K2:K'+MR+'*N(E2:E'+MR+')))');
  setAF_(sh, 'N2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(K2:K'+MR+'<0,"Minus",IF(K2:K'+MR+'=0,"Habis",IF(K2:K'+MR+'<=N(M2:M'+MR+'),"Rendah","Aman")))))');
  setAF_(sh, 'O2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(K2:K'+MR+'<0,"⚠️ Stok minus",IF(K2:K'+MR+'<=N(M2:M'+MR+'),"⚠️ Stok menipis",""))))');
  sh.getRange('P1').setValue('laba_terjual (sys)');
  setAF_(sh, 'P2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+JUAL+'!$E$2:$E$'+MR+',A2:A'+MR+','+JUAL+'!$W$2:$W$'+MR+')))');
  sh.hideColumns(16);
}
function formulasTransaksiKas_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS);
  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","TKS-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(((C2:C'+MR+'="tarik_tunai")+(C2:C'+MR+'="transfer_uang")>0)*(COUNTIF(FINAL_STATUS,K2:K'+MR+')>0),N(H2:H'+MR+'),0)))');
  setAF_(sh, 'P2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF((C2:C'+MR+'="koreksi_kas")*(COUNTIF(FINAL_STATUS,K2:K'+MR+')=0),"⚠️ Koreksi kas belum Approved",'+
    'IF((C2:C'+MR+'="pengeluaran")*(E2:E'+MR+'=""),"⚠️ Akun Keluar kosong",'+
    'IF((C2:C'+MR+'="modal")*(F2:F'+MR+'=""),"⚠️ Akun Masuk kosong",'+
    'IF(((C2:C'+MR+'="mutasi")+(C2:C'+MR+'="tarik_tunai")+(C2:C'+MR+'="transfer_uang")>0)*((E2:E'+MR+'="")+(F2:F'+MR+'="")>0),"⚠️ Akun dari/ke kosong",'+
    'IF(N(G2:G'+MR+')<=0,"⚠️ Nominal kosong/0","")))))))');
  sh.getRange('Q1').setValue('eff_pengeluaran (sys)');
  setAF_(sh, 'Q2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF((C2:C'+MR+'="pengeluaran")*(COUNTIF(FINAL_STATUS,K2:K'+MR+')>0),N(G2:G'+MR+'),0)))');
  sh.hideColumns(17);
}
function formulasKoreksiStok_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.KOREKSI_STOK);
  var P = CFG.SHEET.PRODUK;
  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","KST-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'D2', '=ARRAYFORMULA(IF(C2:C'+MR+'="","",IFERROR(INDEX('+P+'!$A$2:$A$'+MR+',MATCH(C2:C'+MR+','+P+'!$B$2:$B$'+MR+',0)),"❓ tdk ada")))');
  setAF_(sh, 'J2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(C2:C'+MR+'="","⚠️ Produk kosong",'+
    'IF(N(F2:F'+MR+')<=0,"⚠️ Qty kosong/0",'+
    'IF(COUNTIF(FINAL_STATUS,H2:H'+MR+')=0,"⏳ Menunggu Approve","✅ Diterapkan")))))');
  sh.getRange('K1').setValue('eff_koreksi (sys)');
  setAF_(sh, 'K2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(COUNTIF(FINAL_STATUS,H2:H'+MR+')>0,'+
    'IF(E2:E'+MR+'="Tambah",N(F2:F'+MR+'),-N(F2:F'+MR+')),0)))');
  sh.hideColumns(11);
}
function formulasUtangPiutang_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG);
  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","UPI-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'H2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",N(E2:E'+MR+')-N(G2:G'+MR+')))');
  setAF_(sh, 'K2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(H2:H'+MR+'<=0,"✅ Lunas",'+
    'IF(F2:F'+MR+'="","-",'+
    'IF(F2:F'+MR+'<TODAY(),"❗ Telat",'+
    'IF(F2:F'+MR+'<=TODAY()+3,"⚠️ Jatuh Tempo","Aman"))))))');
  setAF_(sh, 'M2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF((H2:H'+MR+'>0)*(F2:F'+MR+'<>"")*(F2:F'+MR+'<TODAY()),"❗ Lewat jatuh tempo","")))');
}
function formulasAkun_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.AKUN);
  var BB = CFG.SHEET.BUKU_BESAR;
  setAF_(sh, 'D2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+BB+'!$E$2:$E$'+MR+',A2:A'+MR+','+BB+'!$F$2:$F$'+MR+')))');
  setAF_(sh, 'E2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+BB+'!$E$2:$E$'+MR+',A2:A'+MR+','+BB+'!$G$2:$G$'+MR+')))');
  setAF_(sh, 'F2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",N(C2:C'+MR+')+D2:D'+MR+'-E2:E'+MR+'))');
  setAF_(sh, 'H2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(G2:G'+MR+'="","",N(G2:G'+MR+')-F2:F'+MR+')))');
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(G2:G'+MR+'="","Belum dicek",IF(ABS(N(G2:G'+MR+')-F2:F'+MR+')<1,"Cocok","Selisih"))))');
  setAF_(sh, 'J2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(F2:F'+MR+'<0,"❗ Saldo negatif",IF((G2:G'+MR+'<>"")*(ABS(N(G2:G'+MR+')-F2:F'+MR+')>=1),"⚠️ Ada selisih",""))))');
}
function applyNumberFormats_(ss) {
  var rp = '"Rp"#,##0;[Red]-"Rp"#,##0';
  var f = function(sheet, cols) {
    var sh = ss.getSheetByName(sheet);
    cols.forEach(function(c){ sh.getRange(c+'2:'+c+MR).setNumberFormat(rp); });
  };
  ss.getSheetByName(CFG.SHEET.PENJUALAN).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
  f(CFG.SHEET.PENJUALAN, ['H','I','J','K','L']);
  ss.getSheetByName(CFG.SHEET.PEMBELIAN).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
  f(CFG.SHEET.PEMBELIAN, ['H','I','N']);
  f(CFG.SHEET.PRODUK, ['E','F','L']);
  ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
  f(CFG.SHEET.TRANSAKSI_KAS, ['G','H','I']);
  ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG).getRange('F2:F'+MR).setNumberFormat('yyyy-mm-dd');
  f(CFG.SHEET.UTANG_PIUTANG, ['E','G','H']);
  f(CFG.SHEET.AKUN, ['C','D','E','F','G','H']);
  ss.getSheetByName(CFG.SHEET.BUKU_BESAR).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
  f(CFG.SHEET.BUKU_BESAR, ['F','G','H']);
  ss.getSheetByName(CFG.SHEET.KOREKSI_STOK).getRange('A2:A'+MR).setNumberFormat('yyyy-mm-dd');
}
function buildAuditDefinitions(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.AUDIT);
  var J = CFG.SHEET.PENJUALAN, B = CFG.SHEET.PEMBELIAN, P = CFG.SHEET.PRODUK,
      T = CFG.SHEET.TRANSAKSI_KAS, K = CFG.SHEET.KOREKSI_STOK, U = CFG.SHEET.UTANG_PIUTANG,
      A = CFG.SHEET.AKUN, BB = CFG.SHEET.BUKU_BESAR;
  var defs = [
    ['Stok minus', '=COUNTIF('+P+'!N2:N'+MR+',"Minus")', 'CRITICAL', 'Lihat sheet produk kolom Status Stok'],
    ['Saldo akun negatif', '=COUNTIF('+A+'!F2:F'+MR+',"<0")', 'CRITICAL', 'Lihat sheet akun kolom Saldo Sistem'],
    ['Penjualan final tanpa Akun Uang Masuk', '=COUNTIFS('+J+'!T2:T'+MR+',"Final",'+J+'!N2:N'+MR+',"")', 'WARNING', 'Lengkapi Akun Uang Masuk'],
    ['Penjualan final tanpa Akun Modal/HPP', '=COUNTIFS('+J+'!T2:T'+MR+',"Final",'+J+'!O2:O'+MR+',"")', 'WARNING', 'Lengkapi Akun Modal / Sumber HPP'],
    ['Produk punya nama tanpa kode', '=COUNTIFS('+P+'!B2:B'+MR+',"?*",'+P+'!A2:A'+MR+',"")', 'WARNING', 'Isi Kode Produk'],
    ['Harga modal kosong', '=COUNTIFS('+P+'!B2:B'+MR+',"?*",'+P+'!E2:E'+MR+',"")', 'WARNING', 'Isi Harga Modal Default'],
    ['Harga jual kosong', '=COUNTIFS('+P+'!B2:B'+MR+',"?*",'+P+'!F2:F'+MR+',"")', 'WARNING', 'Isi Harga Jual Default'],
    ['Penjualan: Qty kosong/0', '=COUNTIFS('+J+'!D2:D'+MR+',"?*",'+J+'!G2:G'+MR+',"")+COUNTIFS('+J+'!D2:D'+MR+',"?*",'+J+'!G2:G'+MR+',0)', 'WARNING', 'Isi Qty > 0'],
    ['Penjualan final data tak lengkap', '=COUNTIF('+J+'!U2:U'+MR+',"?*")', 'WARNING', 'Lihat kolom Warning di penjualan'],
    ['Mutasi tidak balance (akun kosong)', '=COUNTIFS('+T+'!C2:C'+MR+',"mutasi",'+T+'!E2:E'+MR+',"")+COUNTIFS('+T+'!C2:C'+MR+',"mutasi",'+T+'!F2:F'+MR+',"")', 'CRITICAL', 'Mutasi wajib Akun Keluar & Masuk'],
    ['Ledger tanpa ID referensi (anomali)', '=COUNTIFS('+BB+'!E2:E'+MR+',"?*",'+BB+'!B2:B'+MR+',"")', 'CRITICAL', 'Jalankan Refresh Buku Besar'],
    ['Koreksi kas belum Approved', '=COUNTIF('+T+'!P2:P'+MR+',"*belum Approved*")', 'WARNING', 'Approve dulu di transaksi_kas'],
    ['Koreksi stok belum Approved', '=COUNTIFS('+K+'!C2:C'+MR+',"?*",'+K+'!H2:H'+MR+',"Pending")', 'WARNING', 'Approve dulu di koreksi_stok'],
    ['Piutang jatuh tempo / telat', '=COUNTIFS('+U+'!C2:C'+MR+',"Piutang",'+U+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+U+'!C2:C'+MR+',"Piutang",'+U+'!K2:K'+MR+',"*Jatuh Tempo*")', 'WARNING', 'Tagih pelanggan'],
    ['Hutang jatuh tempo / telat', '=COUNTIFS('+U+'!C2:C'+MR+',"Hutang",'+U+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+U+'!C2:C'+MR+',"Hutang",'+U+'!K2:K'+MR+',"*Jatuh Tempo*")', 'WARNING', 'Bayar supplier']
  ];
  for (var i = 0; i < defs.length; i++) {
    var r = i + 2;
    sh.getRange(r, 1).setValue(i + 1);
    sh.getRange(r, 2).setValue(defs[i][0]);
    sh.getRange(r, 3).setFormula(defs[i][1]);
    sh.getRange(r, 4).setFormula('=IF(C'+r+'=0,"✅ OK","'+(defs[i][2]==='CRITICAL'?'🔴 CRITICAL':'🟠 WARNING')+'")');
    sh.getRange(r, 5).setValue(defs[i][3]);
  }
  sh.setColumnWidth(2, 280); sh.setColumnWidth(5, 300);
  var rng = sh.getRange('D2:D'+(defs.length+1));
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('OK')
    .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('WARNING')
    .setBackground('#FEF3C7').setFontColor(CFG.COLOR.WARN).setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('CRITICAL')
    .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);
}
function rebuildLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Menghitung ulang buku besar...', 'Ledger', 4);
  var finalSet = getFinalStatusSet_(ss);
  var balances = {};
  var akSh = ss.getSheetByName(CFG.SHEET.AKUN);
  var akRows = readRows_(akSh, 3);
  akRows.forEach(function(r){ if (r[0] !== '') balances[r[0]] = num_(r[2]); });
  var E = [];   // {d, ref, src, jenis, akun, masuk, keluar, kat, status, ket, user}
  collectPenjualan_(ss, finalSet, E);
  collectPembelian_(ss, finalSet, E);
  collectTransaksiKas_(ss, finalSet, E);
  collectUtangPiutang_(ss, E);
  E.sort(function(a, b){ return (a.d ? a.d.getTime() : 0) - (b.d ? b.d.getTime() : 0); });
  var out = [];
  E.forEach(function(e){
    if (!(e.akun in balances)) balances[e.akun] = 0;
    balances[e.akun] += (e.masuk - e.keluar);
    out.push([e.d, e.ref, e.src, e.jenis, e.akun, e.masuk, e.keluar, balances[e.akun],
              e.kat, e.status, e.ket, e.user]);
  });
  var bb = ss.getSheetByName(CFG.SHEET.BUKU_BESAR);
  bb.getRange(2, 1, CFG.MAX_ROW - 1, 12).clearContent();
  if (out.length > 0) bb.getRange(2, 1, out.length, 12).setValues(out);
  SpreadsheetApp.flush();
  ss.toast('Buku besar diperbarui: ' + out.length + ' baris.', 'Ledger', 5);
}
function collectPenjualan_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.PENJUALAN);
  var rows = readRows_(sh, 21);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], kat=r[5], totHPP=num_(r[9]), totJual=num_(r[10]),
        metode=r[12], akunMasuk=r[13], akunModal=r[14], status=r[15], user=r[18];
    if (tgl==='' || !finalSet[status]) return;
    if (metode !== 'Kredit' && akunMasuk !== '') {
      E.push(entry_(tgl, id, CFG.SHEET.PENJUALAN, 'Penjualan', akunMasuk, totJual, 0, kat, status, 'Omzet penjualan', user));
    }
    if (akunModal !== '' && totHPP > 0) {
      E.push(entry_(tgl, id, CFG.SHEET.PENJUALAN, 'HPP Penjualan', akunModal, 0, totHPP, kat, status, 'Pemakaian modal/HPP', user));
    }
  });
}
function collectPembelian_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.PEMBELIAN);
  var rows = readRows_(sh, 17);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], kat=r[5], total=num_(r[8]), akunBayar=r[10],
        status=r[11], jenisBeli=r[12], user=r[15];
    if (tgl==='' || !finalSet[status]) return;
    if (jenisBeli !== 'Kredit' && akunBayar !== '') {
      E.push(entry_(tgl, id, CFG.SHEET.PEMBELIAN, 'Pembelian Stok', akunBayar, 0, total, kat, status, 'Beli stok', user));
    }
  });
}
function collectTransaksiKas_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS);
  var rows = readRows_(sh, 16);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], jenis=r[2], kat=r[3], akunK=r[4], akunM=r[5],
        nominal=num_(r[6]), fee=num_(r[7]), status=r[10], arah=r[11], pihak=r[12], user=r[14];
    if (tgl==='' || !finalSet[status]) return;
    switch (jenis) {
      case 'pengeluaran':
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Pengeluaran',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'modal':
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Modal Masuk',akunM,nominal,0,kat,status,pihak,user));
        break;
      case 'mutasi':
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Mutasi (keluar)',akunK,0,nominal,kat,status,pihak,user));
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Mutasi (masuk)',akunM,nominal,0,kat,status,pihak,user));
        if (fee > 0 && akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Biaya Admin Mutasi',akunK,0,fee,kat,status,pihak,user));
        break;
      case 'tarik_tunai':
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Tarik Tunai (masuk)',akunM,nominal+fee,0,kat,status,pihak,user));
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Tarik Tunai (tunai keluar)',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'transfer_uang':
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Transfer Uang (tunai masuk)',akunM,nominal+fee,0,kat,status,pihak,user));
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Transfer Uang (keluar)',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'koreksi_kas':
        if (arah === 'Tambah' && akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Koreksi Kas (+)',akunM,nominal,0,kat,status,pihak,user));
        else if (arah === 'Kurang' && akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Koreksi Kas (-)',akunK,0,nominal,kat,status,pihak,user));
        break;
    }
  });
}
function collectUtangPiutang_(ss, E) {
  var sh = ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG);
  var rows = readRows_(sh, 13);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], arah=r[2], pihak=r[3], dibayar=num_(r[6]), akunBayar=r[8];
    if (tgl==='' || dibayar<=0 || akunBayar==='') return;
    if (arah === 'Piutang') {
      E.push(entry_(tgl,id,CFG.SHEET.UTANG_PIUTANG,'Pelunasan Piutang',akunBayar,dibayar,0,'Piutang','-',pihak,''));
    } else if (arah === 'Hutang') {
      E.push(entry_(tgl,id,CFG.SHEET.UTANG_PIUTANG,'Pembayaran Hutang',akunBayar,0,dibayar,'Hutang','-',pihak,''));
    }
  });
}
function entry_(d, ref, src, jenis, akun, masuk, keluar, kat, status, ket, user) {
  return { d:(d instanceof Date ? d : new Date(d)), ref:ref, src:src, jenis:jenis,
           akun:akun, masuk:masuk, keluar:keluar, kat:kat, status:status, ket:ket, user:user };
}
function getFinalStatusSet_(ss) {
  var dm = ss.getSheetByName(CFG.SHEET.DATA_MASTER);
  var rows = readRows_(dm, 2);
  var set = {};
  rows.forEach(function(r){ if (r[0] !== '' && String(r[1]).toLowerCase() === 'ya') set[r[0]] = true; });
  return set;
}
function readRows_(sh, nCol) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, nCol).getValues();
}
function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}
function applyValidation(ss) {
  fillDropdownSource_(ss);
  var DD = CFG.SHEET.DROPDOWN;
  dv_(ss, CFG.SHEET.PENJUALAN, 'C', DD, 'I');   // Jenis Transaksi
  dv_(ss, CFG.SHEET.PENJUALAN, 'D', DD, 'A');   // Nama Produk
  dv_(ss, CFG.SHEET.PENJUALAN, 'M', DD, 'D');   // Metode
  dv_(ss, CFG.SHEET.PENJUALAN, 'N', DD, 'B');   // Akun Uang Masuk
  dv_(ss, CFG.SHEET.PENJUALAN, 'O', DD, 'B');   // Akun Modal / Sumber HPP
  dv_(ss, CFG.SHEET.PENJUALAN, 'P', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.PENJUALAN, 'Q', DD, 'F');   // Pelanggan
  dv_(ss, CFG.SHEET.PENJUALAN, 'S', DD, 'G');   // User
  dv_(ss, CFG.SHEET.PEMBELIAN, 'C', DD, 'E');   // Supplier
  dv_(ss, CFG.SHEET.PEMBELIAN, 'D', DD, 'A');   // Produk
  dv_(ss, CFG.SHEET.PEMBELIAN, 'J', DD, 'D');   // Metode
  dv_(ss, CFG.SHEET.PEMBELIAN, 'K', DD, 'B');   // Akun Pembayaran
  dv_(ss, CFG.SHEET.PEMBELIAN, 'L', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.PEMBELIAN, 'M', DD, 'M');   // Jenis Pembelian
  dv_(ss, CFG.SHEET.PEMBELIAN, 'P', DD, 'G');   // User
  dv_(ss, CFG.SHEET.PRODUK, 'C', DD, 'H');      // Kategori
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'C', DD, 'J');  // Jenis
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'D', DD, 'P');  // Kategori
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'E', DD, 'B');  // Akun Keluar
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'F', DD, 'B');  // Akun Masuk
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'J', DD, 'D');  // Metode
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'K', DD, 'C');  // Status
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'L', DD, 'O');  // Arah Koreksi
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'M', DD, 'F');  // Pihak (saran pelanggan)
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'O', DD, 'G');  // User
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'C', DD, 'A');   // Produk
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'E', DD, 'K');   // Arah (Tambah/Kurang)
  dvDirect_(ss, CFG.SHEET.KOREKSI_STOK, 'G', CFG.SHEET.DATA_MASTER, 'I'); // Alasan = Jenis Koreksi Stok
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'H', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'I', DD, 'G');   // User
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'C', DD, 'L');  // Arah (Piutang/Hutang)
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'D', DD, 'F');  // Pihak
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'I', DD, 'B');  // Akun Bayar
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'J', DD, 'C');  // Status
  dv_(ss, CFG.SHEET.AKUN, 'B', DD, 'Q');           // Jenis Akun
}
function fillDropdownSource_(ss) {
  var dd = ss.getSheetByName(CFG.SHEET.DROPDOWN);
  var DM = CFG.SHEET.DATA_MASTER, PR = CFG.SHEET.PRODUK, AK = CFG.SHEET.AKUN, M = MR;
  var flt = function(rangeA1){ return '=IFERROR(FILTER('+rangeA1+','+rangeA1+'<>""),"")'; };
  dd.getRange('A2').setFormula(flt(PR+'!B2:B'+M));        // Produk (nama)
  dd.getRange('B2').setFormula(flt(AK+'!A2:A'+M));        // Akun
  dd.getRange('C2').setFormula(flt(DM+'!A2:A'+M));        // Status
  dd.getRange('D2').setFormula(flt(DM+'!J2:J'+M));        // Metode
  dd.getRange('E2').setFormula(flt(DM+'!K2:K'+M));        // Supplier
  dd.getRange('F2').setFormula(flt(DM+'!L2:L'+M));        // Pelanggan
  dd.getRange('G2').setFormula(flt(DM+'!M2:M'+M));        // User
  dd.getRange('H2').setFormula(flt(DM+'!C2:C'+M));        // Kategori Produk
  dd.getRange('I2').setFormula(flt(DM+'!E2:E'+M));        // Jenis Jual
  dd.getRange('J2').setFormula(flt(DM+'!F2:F'+M));        // Jenis Kas
  setCol_(dd, 11, ['Tambah','Kurang']);                          // K Arah Stok
  setCol_(dd, 12, ['Piutang','Hutang']);                         // L Arah Utang/Piutang
  setCol_(dd, 13, ['Tunai','Kredit']);                           // M Jenis Pembelian
  setCol_(dd, 14, ['Harian','Mingguan','Bulanan','Tahunan']);    // N Mode Waktu
  setCol_(dd, 15, ['Tambah','Kurang']);                          // O Arah Koreksi Kas
  dd.getRange('P2').setFormula(
    '=IFERROR(FILTER({'+DM+'!D2:D'+M+';'+DM+'!G2:G'+M+';'+DM+'!H2:H'+M+'},'+
    '{'+DM+'!D2:D'+M+';'+DM+'!G2:G'+M+';'+DM+'!H2:H'+M+'}<>""),"")');
  setCol_(dd, 17, ['Kas','Bank','E-Wallet','QRIS','Supplier','Lainnya']);  // Q Jenis Akun
}
function dv_(ss, targetSheet, targetCol, srcSheet, srcCol) {
  var src = ss.getSheetByName(srcSheet).getRange(srcCol + '2:' + srcCol + '1000');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(src, true).setAllowInvalid(true).build();
  ss.getSheetByName(targetSheet).getRange(targetCol + '2:' + targetCol + MR).setDataValidation(rule);
}
function dvDirect_(ss, targetSheet, targetCol, srcSheet, srcCol) {
  dv_(ss, targetSheet, targetCol, srcSheet, srcCol);
}
function buildDashboard(ss) {
  setupDashboardHelper_(ss);
  var sh = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  sh.clear();
  sh.getCharts().forEach(function(c){ sh.removeChart(c); });
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 22);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 96);
  sh.setColumnWidth(14, 22);
  sh.getRange('A1:N72').setBackground(CFG.COLOR.BG_APP);
  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, PR=CFG.SHEET.PRODUK,
      UP=CFG.SHEET.UTANG_PIUTANG, AK=CFG.SHEET.AKUN, K=CFG.SHEET.KALKULASI;
  sh.getRange('B2:M2').merge().setValue('  💼  KEUANGAN KONTER PULSA  •  DASHBOARD')
    .setBackground('#0F172A').setFontColor('#FFFFFF').setFontSize(22).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(2, 58);
  sh.getRange('B3:M3').merge().setFormula(
    '="  📅  "&IF(D7="Semua","Tahun "&B7,TEXT(DATE(B7,IF(D7="Semua",1,D7),1),"mmmm yyyy"))' +
    '&"   ●   📊  "&F7&"   ●   💳  "&H7&"   ●   🏷️  "&K7&"   ●   ✓  "&M7' +
    '&"   ●   🕒  Update: "&TEXT(NOW(),"dd-mmm hh:mm")')
    .setBackground('#1E293B').setFontColor('#CBD5E1').setFontSize(10)
    .setFontStyle('italic').setVerticalAlignment('middle');
  sh.setRowHeight(3, 28);
  sh.setRowHeight(4, 10);
  sh.getRange('B5:M5').merge().setValue('  ⚙️  FILTER  ─  ubah dropdown di bawah, semua angka & grafik akan otomatis menyesuaikan')
    .setBackground(CFG.COLOR.ACCENT).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sh.setRowHeight(5, 22);
  var fLabels = [['TAHUN',2],['BULAN',4],['MODE WAKTU',6],['AKUN',8],['KATEGORI',11],['STATUS',13]];
  fLabels.forEach(function(f){
    sh.getRange(6, f[1]).setValue(f[0]).setFontSize(8).setFontColor(CFG.COLOR.MUTED)
      .setFontWeight('bold').setBackground(CFG.COLOR.BG_APP);
    sh.getRange(7, f[1]).setBackground('#FFFFFF').setFontWeight('bold').setFontColor(CFG.COLOR.ACCENT)
      .setHorizontalAlignment('center').setFontSize(11)
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(6, 16); sh.setRowHeight(7, 30); sh.setRowHeight(8, 14);
  var now = new Date();
  sh.getRange('B7').setValue(now.getFullYear());
  sh.getRange('D7').setValue('Semua');
  sh.getRange('F7').setValue('Bulanan');
  sh.getRange('H7').setValue('Semua');
  sh.getRange('K7').setValue('Semua');
  sh.getRange('M7').setValue('Final');
  filterDV_(ss, sh, 'B7', K, 'AB');
  filterDV_(ss, sh, 'D7', K, 'AC');
  filterDV_(ss, sh, 'F7', CFG.SHEET.DROPDOWN, 'N');
  filterDV_(ss, sh, 'H7', K, 'AD');
  filterDV_(ss, sh, 'K7', K, 'AE');
  filterDV_(ss, sh, 'M7', K, 'AF');
  var dr = '">="&'+K+'!G1', dr2='"<="&'+K+'!G2';
  var omzet = '=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var laba  = '=SUMIFS('+JUAL+'!L2:L'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var beban = '=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  var fee   = 'SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  premiumKPI_(sh, 9,  2, '💰  OMZET',         'Total penjualan periode', omzet,                   CFG.COLOR.ACCENT, K+'!O2:O54');
  premiumKPI_(sh, 9,  5, '📈  LABA KOTOR',    'Jual − HPP',              laba,                    CFG.COLOR.POS,    K+'!P2:P54');
  premiumKPI_(sh, 9,  8, '💸  PENGELUARAN',   'Biaya operasional',       beban,                   CFG.COLOR.NEG,    null);
  premiumKPI_(sh, 9, 11, '🎯  LABA BERSIH',   'Kotor + Admin − Beban',   '=B10+'+fee+'-H10',      CFG.COLOR.POS,    null);
  sh.setRowHeight(13, 14);
  premiumKPI_(sh, 14,  2, '🏦  SALDO KAS/BANK', 'Total saldo akun',     '=SUM('+AK+'!F2:F'+MR+')',                                    CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 14,  5, '📦  NILAI STOK',     'Stok akhir × modal',   '=SUM('+PR+'!L2:L'+MR+')',                                    CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 14,  8, '📥  PIUTANG',        'Tagihan ke pelanggan', '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Piutang")',     CFG.COLOR.WARN,      null);
  premiumKPI_(sh, 14, 11, '📤  HUTANG',         'Kewajiban ke supplier','=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Hutang")',      CFG.COLOR.WARN,      null);
  sh.setRowHeight(18, 14);
  sh.getRange('B19:M19').merge().setValue('  📈  TREN OMZET & LABA  ─  grafik berubah otomatis sesuai Mode Waktu')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(19, 24);
  var kal = ss.getSheetByName(K);
  var chart = sh.newChart().asComboChart()
    .addRange(kal.getRange('N1:P54'))
    .setPosition(20, 2, 4, 4)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'end', textStyle: { fontSize: 11, bold: true } })
    .setOption('width', 1130).setOption('height', 340)
    .setOption('colors', [CFG.COLOR.ACCENT, CFG.COLOR.POS])
    .setOption('seriesType', 'bars')
    .setOption('series', { 0: { labelInLegend: 'Omzet' }, 1: { type: 'line', lineWidth: 3, pointSize: 6, labelInLegend: 'Laba Kotor' } })
    .setOption('hAxis', { textStyle: { fontSize: 10, color: '#475569' } })
    .setOption('vAxis', { format: 'short', textStyle: { fontSize: 10, color: '#475569' }, gridlines: { color: '#E2E8F0' } })
    .setOption('backgroundColor', CFG.COLOR.CARD)
    .setOption('chartArea', { left: 70, top: 30, width: '88%', height: '78%' })
    .setOption('bar', { groupWidth: '60%' })
    .build();
  sh.insertChart(chart);
  for (var rr = 20; rr <= 35; rr++) sh.setRowHeight(rr, 22);
  sh.setRowHeight(36, 14);
  sh.getRange('B37:G37').merge().setValue('  🏆  TOP 5 PRODUK  ─  qty terjual')
    .setBackground(CFG.COLOR.POS).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.getRange('H37:M37').merge().setValue('  🔔  PANEL ALERT  ─  status sistem')
    .setBackground(CFG.COLOR.NEG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(37, 24);
  sh.getRange('B38:D38').merge().setValue('Produk').setBackground('#F1F5F9')
    .setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.getRange('E38').setValue('Qty').setBackground('#F1F5F9')
    .setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED).setHorizontalAlignment('center');
  sh.getRange('F38:G38').merge().setValue('Visualisasi').setBackground('#F1F5F9')
    .setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.setRowHeight(38, 22);
  for (var i = 0; i < 5; i++) {
    var r = 39 + i;
    sh.getRange(r, 2, 1, 3).merge().setFormula(
      '=IFERROR(INDEX(QUERY('+PR+'!A2:O'+MR+',"select B where I > 0 order by I desc limit 5",0),'+(i+1)+',1),"")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(r, 5).setFormula(
      '=IFERROR(INDEX(QUERY('+PR+'!A2:O'+MR+',"select I where I > 0 order by I desc limit 5",0),'+(i+1)+',1),"")')
      .setBackground('#FFFFFF').setHorizontalAlignment('center').setNumberFormat('#,##0')
      .setFontWeight('bold').setFontColor(CFG.COLOR.HEADER_BG);
    sh.getRange(r, 6, 1, 2).merge().setFormula(
      '=IFERROR(IF(MAX($E$39:$E$43)=0,"",REPT("█",ROUND(E'+r+'/MAX($E$39:$E$43)*15,0))),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.POS).setFontSize(11)
      .setFontFamily('Consolas').setVerticalAlignment('middle');
    sh.setRowHeight(r, 24);
  }
  sh.getRange('B38:G43').setBorder(true,true,true,true,true,true,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
  var alerts = [
    ['📦 Stok minus',         '=COUNTIF('+PR+'!N2:N'+MR+',"Minus")'],
    ['🟡 Stok habis',         '=COUNTIF('+PR+'!N2:N'+MR+',"Habis")'],
    ['🟠 Stok rendah',        '=COUNTIF('+PR+'!N2:N'+MR+',"Rendah")'],
    ['💸 Saldo akun negatif', '=COUNTIF('+AK+'!F2:F'+MR+',"<0")'],
    ['📥 Piutang jatuh tempo','=COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['📤 Hutang jatuh tempo', '=COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['⏳ Transaksi non-final','=COUNTIF('+JUAL+'!P2:P'+MR+',"Pending")+COUNTIF('+JUAL+'!P2:P'+MR+',"Draft")+COUNTIF('+JUAL+'!P2:P'+MR+',"Proses")']
  ];
  for (var i = 0; i < alerts.length; i++) {
    var r = 38 + i;
    sh.getRange(r, 8, 1, 5).merge().setValue(alerts[i][0]).setFontSize(10)
      .setBackground('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true,true,true,true,false,false,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 13).setFormula(alerts[i][1]).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(11)
      .setBackground('#FFFFFF')
      .setBorder(true,true,true,true,false,false,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r, 24);
  }
  var aRange = sh.getRange(38, 13, alerts.length, 1);
  var existing = sh.getConditionalFormatRules();
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([aRange]).build());
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0)
    .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([aRange]).build());
  sh.setConditionalFormatRules(existing);
  sh.setRowHeight(45, 14);
  sh.getRange('B46:G46').merge().setValue('  🏦  SALDO PER AKUN  ─  dengan persentase')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.getRange('H46:M46').merge().setValue('  💸  TOP 5 PENGELUARAN  ─  per kategori')
    .setBackground(CFG.COLOR.NEG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(46, 24);
  sh.getRange('B47:C47').merge().setValue('Akun').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.getRange('D47').setValue('Jenis').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.getRange('E47').setValue('Saldo').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED).setHorizontalAlignment('right');
  sh.getRange('F47:G47').merge().setValue('Distribusi').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.setRowHeight(47, 22);
  for (var i = 0; i < 8; i++) {
    var r = 48 + i;
    sh.getRange(r, 2, 1, 2).merge().setFormula('=IFERROR(IF(INDEX('+AK+'!A2:A,'+(i+1)+')="","",INDEX('+AK+'!A2:A,'+(i+1)+')),"")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(r, 4).setFormula('=IFERROR(INDEX('+AK+'!B2:B,'+(i+1)+'),"")')
      .setBackground('#FFFFFF').setFontSize(9).setFontColor(CFG.COLOR.MUTED).setVerticalAlignment('middle');
    sh.getRange(r, 5).setFormula('=IFERROR(IF(INDEX('+AK+'!A2:A,'+(i+1)+')="","",INDEX('+AK+'!F2:F,'+(i+1)+')),"")')
      .setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0').setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange(r, 6, 1, 2).merge().setFormula(
      '=IFERROR(IF(OR(E'+r+'="",E'+r+'<=0,SUM('+AK+'!F2:F'+MR+')<=0),"",REPT("█",ROUND(E'+r+'/SUM('+AK+'!F2:F'+MR+')*15,0))&"  "&TEXT(E'+r+'/SUM('+AK+'!F2:F'+MR+'),"0%")),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.ACCENT).setFontSize(10)
      .setFontFamily('Consolas').setVerticalAlignment('middle');
    sh.setRowHeight(r, 22);
  }
  sh.getRange('B47:G55').setBorder(true,true,true,true,true,true,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('H47:I47').merge().setValue('Kategori').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  sh.getRange('J47').setValue('Nominal').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED).setHorizontalAlignment('right');
  sh.getRange('K47:M47').merge().setValue('Visualisasi').setBackground('#F1F5F9').setFontWeight('bold').setFontSize(9).setFontColor(CFG.COLOR.MUTED);
  for (var i = 0; i < 5; i++) {
    var r = 48 + i;
    sh.getRange(r, 8, 1, 2).merge().setFormula(
      '=IFERROR(INDEX(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),'+(i+1)+',1),"")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(r, 10).setFormula(
      '=IFERROR(INDEX(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),'+(i+1)+',2),"")')
      .setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0').setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange(r, 11, 1, 3).merge().setFormula(
      '=IFERROR(IF(OR(J'+r+'="",J'+r+'<=0,MAX($J$48:$J$52)<=0),"",REPT("█",ROUND(J'+r+'/MAX($J$48:$J$52)*18,0))),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.NEG).setFontSize(11)
      .setFontFamily('Consolas').setVerticalAlignment('middle');
    sh.setRowHeight(r, 22);
  }
  sh.getRange('H47:M52').setBorder(true,true,true,true,true,true,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(56, 12);
  sh.getRange('B57:M57').merge().setFormula(
    '="📝  Catatan Audit: "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*CRITICAL*")&" kritis, "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*WARNING*")&" peringatan.   ●   ⚙️ Klik menu Keuangan → 🔄 Refresh setelah input transaksi.   ●   Filter di atas mengubah seluruh angka dashboard."')
    .setFontColor('#475569').setFontSize(9).setFontStyle('italic')
    .setBackground('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(57, 28);
}
function premiumKPI_(sh, r, c, label, subtitle, formulaValue, accent, sparklineRange) {
  sh.getRange(r, c, 1, 3).merge().setValue(label).setFontSize(10).setFontColor('#475569')
    .setFontWeight('bold').setBackground('#FFFFFF').setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sh.getRange(r+1, c, 1, 3).merge().setFormula(formulaValue).setFontSize(18).setFontWeight('bold')
    .setFontColor(accent).setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r+2, c, 1, 3).merge().setValue(subtitle).setFontSize(8).setFontColor(CFG.COLOR.MUTED)
    .setFontStyle('italic').setBackground('#FFFFFF').setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  if (sparklineRange) {
    sh.getRange(r+3, c, 1, 3).merge().setFormula(
      '=SPARKLINE(' + sparklineRange + ', {"charttype","column";"color","' + accent + '";"empty","zero";"nan","convert"})')
      .setBackground('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else {
    sh.getRange(r+3, c, 1, 3).merge().setValue('').setBackground('#FFFFFF');
  }
  sh.getRange(r, c, 4, 3).setBorder(true,true,true,true,false,false,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 18); sh.setRowHeight(r+1, 38); sh.setRowHeight(r+2, 18); sh.setRowHeight(r+3, 32);
}
function sectionTitle_(sh, r, c, title) {
  sh.getRange(r, c, 1, 3).merge().setValue(title).setFontWeight('bold')
    .setFontColor(CFG.COLOR.HEADER_BG).setFontSize(10).setBackground('#E8EEF5');
}
function filterDV_(ss, sh, cell, srcSheet, srcCol) {
  var src = ss.getSheetByName(srcSheet).getRange(srcCol + '2:' + srcCol + '60');
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(src, true)
    .setAllowInvalid(true).build();
  sh.getRange(cell).setDataValidation(rule);
}
function setupDashboardHelper_(ss) {
  var k = ss.getSheetByName(CFG.SHEET.KALKULASI);
  if (k.getMaxColumns() < 40) k.insertColumnsAfter(k.getMaxColumns(), 40 - k.getMaxColumns());
  var D = CFG.SHEET.DASHBOARD, JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS,
      PR=CFG.SHEET.PRODUK, DM=CFG.SHEET.DATA_MASTER, AK=CFG.SHEET.AKUN;
  k.getRange('D1').setValue('tahun');  k.getRange('E1').setFormula('='+D+'!B7');
  k.getRange('D2').setValue('bulan');  k.getRange('E2').setFormula('='+D+'!D7');
  k.getRange('D3').setValue('mode');   k.getRange('E3').setFormula('='+D+'!F7');
  k.getRange('D4').setValue('akun');   k.getRange('E4').setFormula('='+D+'!H7');
  k.getRange('D5').setValue('kategori');k.getRange('E5').setFormula('='+D+'!K7');
  k.getRange('E8').setFormula('=IF(E5="Semua","*",E5)');   // kat criteria
  k.getRange('E9').setFormula('=IF(E4="Semua","*",E4)');   // akun criteria
  k.getRange('F1').setValue('startDate'); k.getRange('G1').setFormula('=IF(E2="Semua",DATE(E1,1,1),DATE(E1,E2,1))');
  k.getRange('F2').setValue('endDate');   k.getRange('G2').setFormula('=IF(E2="Semua",DATE(E1,12,31),EOMONTH(DATE(E1,E2,1),0))');
  var idx = [];
  for (var i = 1; i <= 53; i++) idx.push([i]);
  k.getRange(2, 11, 53, 1).setValues(idx);   // K2:K54
  k.getRange('N1').setValue('Periode'); k.getRange('O1').setValue('Omzet'); k.getRange('P1').setValue('Laba');
  var bln = 'IF($E$2="Semua",1,$E$2)';
  k.getRange('L2').setFormula(
    '=ARRAYFORMULA(IF(K2:K54="","",'+
    'IF($E$3="Bulanan",IF(K2:K54<=12,DATE($E$1,K2:K54,1),""),'+
    'IF($E$3="Tahunan",IF(K2:K54<=7,DATE($E$1-7+K2:K54,1,1),""),'+
    'IF($E$3="Harian",IF(K2:K54<=DAY(EOMONTH(DATE($E$1,'+bln+',1),0)),DATE($E$1,'+bln+',K2:K54),""),'+
    'IF($E$3="Mingguan",IF(K2:K54<=53,DATE($E$1,1,1)+(K2:K54-1)*7,""),""))))))');
  k.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(L2:L54="","",'+
    'IF($E$3="Bulanan",EOMONTH(L2:L54,0),'+
    'IF($E$3="Tahunan",DATE(YEAR(L2:L54),12,31),'+
    'IF($E$3="Harian",L2:L54,IF($E$3="Mingguan",L2:L54+6,L2:L54))))))');
  k.getRange('N2').setFormula(
    '=ARRAYFORMULA(IF(L2:L54="","",'+
    'IF($E$3="Bulanan",TEXT(L2:L54,"mmm yy"),'+
    'IF($E$3="Tahunan",TEXT(YEAR(L2:L54),"0"),'+
    'IF($E$3="Harian",TEXT(L2:L54,"d/m"),IF($E$3="Mingguan","Mg"&K2:K54,""))))))');
  var fO = [], fP = [];
  for (var r = 2; r <= 54; r++) {
    fO.push(['=IF(L'+r+'="","",SUMIFS('+JUAL+'!$K$2:$K$'+MR+','+JUAL+'!$T$2:$T$'+MR+',"Final",'+JUAL+'!$A$2:$A$'+MR+',">="&L'+r+','+JUAL+'!$A$2:$A$'+MR+',"<="&M'+r+','+JUAL+'!$F$2:$F$'+MR+',$E$8))']);
    fP.push(['=IF(L'+r+'="","",SUMIFS('+JUAL+'!$L$2:$L$'+MR+','+JUAL+'!$T$2:$T$'+MR+',"Final",'+JUAL+'!$A$2:$A$'+MR+',">="&L'+r+','+JUAL+'!$A$2:$A$'+MR+',"<="&M'+r+','+JUAL+'!$F$2:$F$'+MR+',$E$8))']);
  }
  k.getRange(2, 15, 53, 1).setFormulas(fO);   // O2:O54
  k.getRange(2, 16, 53, 1).setFormulas(fP);   // P2:P54
  var yr = new Date().getFullYear();
  var years = [];
  for (var y = yr - 5; y <= yr + 1; y++) years.push([y]);
  k.getRange(2, 28, years.length, 1).setValues(years);                 // AB
  var blnList = [['Semua'],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12]];
  k.getRange(2, 29, blnList.length, 1).setValues(blnList);             // AC
  k.getRange('AD2').setFormula('=IFERROR({"Semua";FILTER('+AK+'!A2:A'+MR+','+AK+'!A2:A'+MR+'<>"")},"Semua")');  // AD
  k.getRange('AE2').setFormula('=IFERROR({"Semua";FILTER('+DM+'!C2:C'+MR+','+DM+'!C2:C'+MR+'<>"")},"Semua")');  // AE
  k.getRange(2, 32, 3, 1).setValues([['Semua'],['Final'],['Non-Final']]); // AF
}
function buildLaporan(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.LAPORAN);
  sh.clear();
  sh.setHiddenGridlines(true);
  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, K=CFG.SHEET.KALKULASI;
  sh.getRange('A1:G1').merge().setValue('LAPORAN LABA RUGI (per Bulan)')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(14).setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  sh.getRange('A2').setValue('Tahun:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B2').setValue(new Date().getFullYear()).setFontWeight('bold')
    .setFontColor(CFG.COLOR.ACCENT).setHorizontalAlignment('center')
    .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  var yrRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(K).getRange('AB2:AB60'), true).setAllowInvalid(true).build();
  sh.getRange('B2').setDataValidation(yrRule);
  var head = ['Periode','Omzet','HPP','Laba Kotor','Admin Fee','Pengeluaran','Laba Bersih'];
  sh.getRange(4, 1, 1, 7).setValues([head]).setBackground(CFG.COLOR.HEADER_BG)
    .setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  var months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  for (var m = 1; m <= 12; m++) {
    var r = 4 + m;
    var s = 'DATE($B$2,'+m+',1)', e = 'EOMONTH(DATE($B$2,'+m+',1),0)';
    sh.getRange(r, 1).setValue(months[m-1]);
    sh.getRange(r, 2).setFormula('=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+',">="&'+s+','+JUAL+'!A2:A'+MR+',"<="&'+e+')');
    sh.getRange(r, 3).setFormula('=SUMIFS('+JUAL+'!J2:J'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+',">="&'+s+','+JUAL+'!A2:A'+MR+',"<="&'+e+')');
    sh.getRange(r, 4).setFormula('=B'+r+'-C'+r);
    sh.getRange(r, 5).setFormula('=SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+',">="&'+s+','+TK+'!A2:A'+MR+',"<="&'+e+')');
    sh.getRange(r, 6).setFormula('=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+',">="&'+s+','+TK+'!A2:A'+MR+',"<="&'+e+')');
    sh.getRange(r, 7).setFormula('=D'+r+'+E'+r+'-F'+r);
  }
  var tr = 17;
  sh.getRange(tr, 1).setValue('TOTAL').setFontWeight('bold');
  for (var col = 2; col <= 7; col++) {
    sh.getRange(tr, col).setFormula('=SUM('+colLetter_(col)+'5:'+colLetter_(col)+'16)').setFontWeight('bold');
  }
  sh.getRange('B5:G17').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0');
  sh.getRange(tr, 1, 1, 7).setBackground('#E8EEF5');
  for (var w = 1; w <= 7; w++) sh.setColumnWidth(w, w === 1 ? 110 : 120);
}
var FORMULA_COLS = {
  penjualan:     ['B','E','F','J','K','L','T','U','V','W'],
  pembelian:     ['B','E','F','I','N','Q','R'],
  produk:        ['H','I','J','K','L','N','O'],
  transaksi_kas: ['B','I','P','Q'],
  koreksi_stok:  ['B','D','J','K'],
  utang_piutang: ['B','H','K','M'],
  akun:          ['D','E','F','H','I','J'],
  skenario_uji:  ['D','E']
};
var FULL_PROTECT = ['buku_besar','audit','_kalkulasi','_dropdown'];
function applyProtection(ss) {
  Object.keys(FORMULA_COLS).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    FORMULA_COLS[name].forEach(function(col){
      var rng = sh.getRange(col + '2:' + col + MR);
      var p = rng.protect().setDescription('Kolom formula - jangan diubah manual');
      p.setWarningOnly(true);
    });
  });
  FULL_PROTECT.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.protect().setDescription('Sheet sistem/hasil - otomatis').setWarningOnly(true);
  });
  var lap = ss.getSheetByName(CFG.SHEET.LAPORAN);
  if (lap) lap.getRange('A4:G17').protect().setDescription('Tabel laba rugi - otomatis').setWarningOnly(true);
  var dash = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  if (dash) {
    dash.getRange('B2:M5').protect().setDescription('Header dashboard').setWarningOnly(true);
    dash.getRange('B6:M6').protect().setDescription('Label filter').setWarningOnly(true);
    dash.getRange('B9:M58').protect().setDescription('Area dashboard - otomatis').setWarningOnly(true);
  }
  logChange(ss, 'SISTEM', 'PROTEKSI', 'Proteksi warning-only diterapkan');
}
function removeAllProtections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sh){
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){ p.remove(); });
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Semua proteksi dilepas.', 'Proteksi', 4);
}
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET.PENJUALAN) return;
  var startCol = e.range.getColumn();
  var endCol = startCol + e.range.getNumColumns() - 1;
  if (4 < startCol || 4 > endCol) return;
  var startRow = e.range.getRow();
  var nRows = e.range.getNumRows();
  if (startRow < 2) startRow = 2;
  var ss = e.source;
  var prMap = buildProdukPriceMap_(ss);
  for (var i = 0; i < nRows; i++) {
    var row = e.range.getRow() + i;
    if (row < 2) continue;
    freezeSnapshotRow_(sh, prMap, row);
  }
}
function freezeSnapshotRow_(sh, prMap, row) {
  var nama = String(sh.getRange(row, 4).getValue()).trim();   // D
  if (nama === '') {
    sh.getRange(row, 8, 1, 2).clearContent();                 // H,I
    return;
  }
  var p = prMap[nama];
  if (!p) return;                                             // produk tak dikenal -> biarkan
  sh.getRange(row, 8).setValue(p.modal);                      // H Harga Modal Snapshot
  sh.getRange(row, 9).setValue(p.jual);                       // I Harga Jual Snapshot
}
function buildProdukPriceMap_(ss) {
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);
  var last = pr.getLastRow();
  var map = {};
  if (last < 2) return map;
  var vals = pr.getRange(2, 2, last - 1, 5).getValues();      // B..F : nama,kategori,satuan,modal,jual
  vals.forEach(function(r){
    var nama = String(r[0]).trim();
    if (nama !== '') map[nama] = { modal: num_(r[3]), jual: num_(r[4]) };
  });
  return map;
}
function freezeAllSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET.PENJUALAN);
  var last = sh.getLastRow();
  if (last < 2) return;
  var prMap = buildProdukPriceMap_(ss);
  var data = sh.getRange(2, 4, last - 1, 6).getValues();      // D..I
  for (var i = 0; i < data.length; i++) {
    var nama = String(data[i][0]).trim();
    var hModal = data[i][4], hJual = data[i][5];               // H,I
    if (nama !== '' && (hModal === '' || hJual === '')) {
      freezeSnapshotRow_(sh, prMap, i + 2);
    }
  }
}
function insertDummyData(ss) {
  var now = new Date();
  var Y = now.getFullYear();
  var cm = now.getMonth() + 1;
  var months = [];
  for (var mm = Math.max(1, cm - 2); mm <= cm; mm++) months.push(mm);
  var prMap = buildProdukPriceMap_(ss);
  var sales = [];
  months.forEach(function(m){
    sales.push(mkSale_(Y,m,5,'Pulsa 10K','Pulsa',5,'QRIS','QRIS','Saldo Supplier','Selesai','Umum','Kasir 1',prMap));
    sales.push(mkSale_(Y,m,8,'Paket Data 1GB','Paket Data',3,'Tunai','Kas Utama','Saldo Supplier','Selesai','Budi','Kasir 1',prMap));
    sales.push(mkSale_(Y,m,12,'Token Listrik 20K','Token Listrik',2,'Transfer Bank','Bank BCA','Saldo Supplier','Lunas','Siti','Owner',prMap));
    sales.push(mkSale_(Y,m,18,'Pulsa 25K','Pulsa',2,'E-Wallet','GoPay','Saldo Supplier','Sukses','Umum','Kasir 2',prMap));
  });
  sales.push(mkSale_(Y,cm,20,'Voucher Game 10K','Voucher Game',4,'Tunai','Kas Utama','Saldo Supplier','Pending','Andi','Kasir 1',prMap)); // pending: tidak dihitung
  sales.push(mkSale_(Y,cm,22,'Paket Data 5GB','Paket Data',2,'Kredit','Kas Utama','Saldo Supplier','Selesai','Toko Maju','Owner',prMap)); // kredit -> piutang
  writeDummy_(ss.getSheetByName(CFG.SHEET.PENJUALAN), sales,
    {A:'tgl',C:'jenis',D:'produk',G:'qty',H:'modal',I:'jual',M:'metode',N:'akunMasuk',O:'akunModal',P:'status',Q:'pel',S:'user'});
  var buys = [
    mkBuy_(Y,months[0],3,'Supplier Pulsa A','Pulsa 10K',50,10500,'Transfer Bank','Bank BCA','Selesai','Tunai','Topup awal','Owner'),
    mkBuy_(Y,months[months.length-1],4,'Distributor Data B','Paket Data 1GB',40,9000,'Transfer Bank','Bank BCA','Selesai','Tunai','Stok data','Owner'),
    mkBuy_(Y,cm,2,'Agen Token D','Token Listrik 20K',30,20500,'Tunai','Kas Utama','Lunas','Tunai','Stok token','Kasir 1'),
    mkBuy_(Y,cm,6,'PPOB Center C','Voucher Game 10K',20,9500,'Kredit','Saldo Supplier','Selesai','Kredit','Beli kredit','Owner'),
    mkBuy_(Y,cm,7,'Supplier Pulsa A','Pulsa 25K',20,24800,'Transfer Bank','Bank BCA','Pending','Tunai','Belum diterima','Owner') // pending: tidak nambah stok
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.PEMBELIAN), buys,
    {A:'tgl',C:'supplier',D:'produk',G:'qty',H:'harga',J:'metode',K:'akunBayar',L:'status',M:'jenisBeli',O:'catatan',P:'user'});
  var tk = [
    mkTk_(Y,cm,1,'pengeluaran','Listrik','Kas Utama','',150000,0,'Tunai','Lunas','','PLN','Bayar listrik','Owner'),
    mkTk_(Y,cm,1,'modal','','','Bank BCA',2000000,0,'Transfer Bank','Selesai','','Pemilik','Setor modal','Owner'),
    mkTk_(Y,cm,3,'mutasi','Setor ke Bank','Kas Utama','Bank BCA',500000,0,'Transfer Bank','Selesai','','','Setor kas ke bank','Owner'),
    mkTk_(Y,cm,4,'tarik_tunai','','Kas Utama','Bank BCA',300000,5000,'Transfer Bank','Sukses','','Rina','Layanan tarik tunai','Kasir 1'),
    mkTk_(Y,cm,5,'transfer_uang','','Bank BCA','Kas Utama',200000,4000,'Tunai','Sukses','','Andi','Layanan transfer','Kasir 1'),
    mkTk_(Y,cm,6,'koreksi_kas','Selisih Lebih','','Kas Utama',10000,0,'Tunai','Approved','Tambah','','Koreksi selisih kas','Owner'),
    mkTk_(Y,cm,2,'pengeluaran','Konsumsi','Kas Utama','',25000,0,'Tunai','Pending','','','Snack (pending - diabaikan)','Kasir 1')
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS), tk,
    {A:'tgl',C:'jenis',D:'kategori',E:'akunKeluar',F:'akunMasuk',G:'nominal',H:'fee',J:'metode',K:'status',L:'arah',M:'pihak',N:'keterangan',O:'user'});
  var ks = [
    mkKs_(Y,cm,5,'Voucher Game 10K','Kurang',2,'Barang Rusak','Approved','Owner'),
    mkKs_(Y,cm,6,'Pulsa 10K','Tambah',3,'Stok Opname','Pending','Kasir 1') // pending: tidak diterapkan
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.KOREKSI_STOK), ks,
    {A:'tgl',C:'produk',E:'arah',F:'qty',G:'alasan',H:'status',I:'user'});
  var up = [
    mkUp_(Y,cm,2,'Piutang','Toko Maju',64000,dt_(Y,cm,16),0,'','Proses','Penjualan kredit Paket Data 5GB'),
    mkUp_(Y,cm,3,'Hutang','PPOB Center C',190000,dt_(Y,cm,20),0,'','Proses','Pembelian voucher kredit'),
    mkUp_(Y,months[0],10,'Piutang','Warung Bu Tini',50000,dt_(Y,months[0],15),20000,'Kas Utama','Proses','Piutang lama (telat)')
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG), up,
    {A:'tgl',C:'arah',D:'pihak',E:'nominal',F:'jatuhTempo',G:'dibayar',I:'akunBayar',J:'status',L:'keterangan'});
  SpreadsheetApp.flush();
}
function dt_(y,m,d){ return new Date(y, m-1, d); }
function mkSale_(y,m,d,produk,jenis,qty,metode,akunMasuk,akunModal,status,pel,user,prMap){
  var p = prMap[produk] || {modal:0,jual:0};
  return {tgl:dt_(y,m,d),jenis:jenis,produk:produk,qty:qty,modal:p.modal,jual:p.jual,
    metode:metode,akunMasuk:akunMasuk,akunModal:akunModal,status:status,pel:pel,user:user};
}
function mkBuy_(y,m,d,supplier,produk,qty,harga,metode,akunBayar,status,jenisBeli,catatan,user){
  return {tgl:dt_(y,m,d),supplier:supplier,produk:produk,qty:qty,harga:harga,metode:metode,
    akunBayar:akunBayar,status:status,jenisBeli:jenisBeli,catatan:catatan,user:user};
}
function mkTk_(y,m,d,jenis,kategori,akunKeluar,akunMasuk,nominal,fee,metode,status,arah,pihak,ket,user){
  return {tgl:dt_(y,m,d),jenis:jenis,kategori:kategori,akunKeluar:akunKeluar,akunMasuk:akunMasuk,
    nominal:nominal,fee:fee,metode:metode,status:status,arah:arah,pihak:pihak,keterangan:ket,user:user};
}
function mkKs_(y,m,d,produk,arah,qty,alasan,status,user){
  return {tgl:dt_(y,m,d),produk:produk,arah:arah,qty:qty,alasan:alasan,status:status,user:user};
}
function mkUp_(y,m,d,arah,pihak,nominal,jt,dibayar,akunBayar,status,ket){
  return {tgl:dt_(y,m,d),arah:arah,pihak:pihak,nominal:nominal,jatuhTempo:jt,dibayar:dibayar,
    akunBayar:akunBayar,status:status,keterangan:ket};
}
function writeDummy_(sh, rows, colMap) {
  var n = rows.length;
  Object.keys(colMap).forEach(function(col){
    var cnum = letterToNum_(col);
    sh.getRange(2, cnum, CFG.MAX_ROW - 1, 1).clearContent();
    if (n > 0) {
      var field = colMap[col];
      var arr = rows.map(function(o){ var v = o[field]; return [ (v === undefined ? '' : v) ]; });
      sh.getRange(2, cnum, n, 1).setValues(arr);
    }
  });
}
function letterToNum_(s){ var n=0; for (var i=0;i<s.length;i++) n = n*26 + (s.charCodeAt(i)-64); return n; }
function seedSkenarioUji(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.SKENARIO);
  var J=CFG.SHEET.PENJUALAN,B=CFG.SHEET.PEMBELIAN,P=CFG.SHEET.PRODUK,T=CFG.SHEET.TRANSAKSI_KAS,
      U=CFG.SHEET.UTANG_PIUTANG,A=CFG.SHEET.AKUN,BB=CFG.SHEET.BUKU_BESAR,L=CFG.SHEET.LAPORAN;
  var defs = [
    ['Penjualan final mengurangi stok','Stok keluar > 0','=SUM('+P+'!I2:I'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Penjualan Pending TIDAK kurangi stok','eff_keluar pending = 0','=SUMIF('+J+'!P2:P'+MR+',"Pending",'+J+'!V2:V'+MR+')','=IF(D{r}=0,"✅ PASS","❌ FAIL")'],
    ['Pembelian final menambah stok','Stok masuk > 0','=SUM('+P+'!H2:H'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Pembelian Pending TIDAK menambah stok','eff_masuk pending = 0','=SUMIF('+B+'!L2:L'+MR+',"Pending",'+B+'!R2:R'+MR+')','=IF(D{r}=0,"✅ PASS","❌ FAIL")'],
    ['Pengeluaran final mengurangi kas','Pengeluaran di ledger > 0','=SUMIF('+BB+'!D2:D'+MR+',"Pengeluaran",'+BB+'!G2:G'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Modal final menambah kas','Modal di ledger > 0','=SUMIF('+BB+'!D2:D'+MR+',"Modal Masuk",'+BB+'!F2:F'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Mutasi balance (masuk = keluar)','Selisih 0','=SUMIF('+BB+'!D2:D'+MR+',"Mutasi (masuk)",'+BB+'!F2:F'+MR+')-SUMIF('+BB+'!D2:D'+MR+',"Mutasi (keluar)",'+BB+'!G2:G'+MR+')','=IF(ROUND(D{r},0)=0,"✅ PASS","❌ FAIL")'],
    ['Tarik tunai: hanya admin fee jadi laba','Laba admin tarik tunai > 0','=SUMIF('+T+'!C2:C'+MR+',"tarik_tunai",'+T+'!I2:I'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Transfer uang: hanya admin fee jadi laba','Laba admin transfer > 0','=SUMIF('+T+'!C2:C'+MR+',"transfer_uang",'+T+'!I2:I'+MR+')','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Koreksi kas Approved ubah saldo','Ada baris koreksi kas di ledger','=COUNTIF('+BB+'!D2:D'+MR+',"Koreksi Kas*")','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Koreksi stok Approved ubah stok','Total koreksi <> 0','=SUM('+P+'!J2:J'+MR+')','=IF(D{r}<>0,"✅ PASS","❌ FAIL")'],
    ['Piutang jatuh tempo muncul warning','Ada piutang telat/JT','=COUNTIFS('+U+'!C2:C'+MR+',"Piutang",'+U+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+U+'!C2:C'+MR+',"Piutang",'+U+'!K2:K'+MR+',"*Jatuh Tempo*")','=IF(D{r}>0,"✅ PASS","❌ FAIL")'],
    ['Hutang jatuh tempo terdeteksi','Hitung hutang telat/JT (info)','=COUNTIFS('+U+'!C2:C'+MR+',"Hutang",'+U+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+U+'!C2:C'+MR+',"Hutang",'+U+'!K2:K'+MR+',"*Jatuh Tempo*")','=IF(D{r}>=0,"✅ OK","❌ FAIL")'],
    ['Stok minus terdeteksi sistem','Hitung stok minus (info)','=COUNTIF('+P+'!N2:N'+MR+',"Minus")','=IF(D{r}>=0,"✅ OK","❌ FAIL")'],
    ['Dashboard ikut filter Bulan/Tahun','Cek manual','="ubah filter Tahun/Bulan di dashboard"','="ℹ️ MANUAL"'],
    ['Dashboard ikut Mode Waktu','Cek manual','="ubah dropdown Mode Waktu di dashboard"','="ℹ️ MANUAL"'],
    ['Ledger TIDAK hitung Batal/Gagal','0 baris status batal/gagal','=COUNTIF('+BB+'!J2:J'+MR+',"Batal")+COUNTIF('+BB+'!J2:J'+MR+',"Gagal")','=IF(D{r}=0,"✅ PASS","❌ FAIL")'],
    ['Laba bersih = LK + admin fee - pengeluaran','Selisih total laporan 0','=ROUND('+L+'!G17-('+L+'!D17+'+L+'!E17-'+L+'!F17),0)','=IF(D{r}=0,"✅ PASS","❌ FAIL")'],
    ['Saldo akun = saldo awal + ledger','Selisih 0','=ROUND(SUM('+A+'!F2:F'+MR+')-(SUM('+A+'!C2:C'+MR+')+SUM('+BB+'!F2:F'+MR+')-SUM('+BB+'!G2:G'+MR+')),0)','=IF(D{r}=0,"✅ PASS","❌ FAIL")'],
    ['Nilai stok = stok akhir x modal','Selisih 0','=ROUND(SUM('+P+'!L2:L100)-SUMPRODUCT(N('+P+'!K2:K100),N('+P+'!E2:E100)),0)','=IF(D{r}=0,"✅ PASS","❌ FAIL")']
  ];
  for (var i = 0; i < defs.length; i++) {
    var r = i + 2;
    sh.getRange(r, 1).setValue(i + 1);
    sh.getRange(r, 2).setValue(defs[i][0]);
    sh.getRange(r, 3).setValue(defs[i][1]);
    sh.getRange(r, 4).setFormula(defs[i][2]);
    sh.getRange(r, 5).setFormula(defs[i][3].replace(/\{r\}/g, r));
  }
  sh.setColumnWidth(2, 290); sh.setColumnWidth(3, 240); sh.setColumnWidth(5, 110);
  var vRange = sh.getRange('E2:E' + (defs.length + 1));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('PASS')
      .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([vRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('FAIL')
      .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([vRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('MANUAL')
      .setBackground('#E2E8F0').setFontColor(CFG.COLOR.MUTED).setRanges([vRange]).build()
  ]);
}
function runTestScenarios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  rebuildLedger();
  SpreadsheetApp.flush();
  var sh = ss.getSheetByName(CFG.SHEET.SKENARIO);
  var last = sh.getLastRow();
  var pass = 0, fail = 0;
  if (last >= 2) {
    var v = sh.getRange(2, 5, last - 1, 1).getValues();
    v.forEach(function(x){ var s=String(x[0]); if (s.indexOf('PASS')>=0||s.indexOf('OK')>=0) pass++; else if (s.indexOf('FAIL')>=0) fail++; });
  }
  ss.setActiveSheet(sh);
  ss.toast('Hasil uji: ' + pass + ' lulus, ' + fail + ' gagal.', 'Skenario Uji', 8);
}
