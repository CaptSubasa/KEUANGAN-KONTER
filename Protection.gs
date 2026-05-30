/**
 * ============================================================================
 *  Protection.gs  -  Proteksi kolom formula (warning-only, tidak memblok input)
 * ----------------------------------------------------------------------------
 *  Memakai setWarningOnly(true) -> user tetap bisa input,
 *  tapi diingatkan kalau menimpa kolom formula / sheet sistem.
 * ============================================================================
 */

// Kolom formula per sheet (yang diproteksi)
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

// Sheet yang diproteksi seluruhnya (warning-only)
var FULL_PROTECT = ['buku_besar','audit','_kalkulasi','_dropdown'];

function applyProtection(ss) {
  // 1. proteksi kolom formula
  Object.keys(FORMULA_COLS).forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    FORMULA_COLS[name].forEach(function(col){
      var rng = sh.getRange(col + '2:' + col + MR);
      var p = rng.protect().setDescription('Kolom formula - jangan diubah manual');
      p.setWarningOnly(true);
    });
  });

  // 2. proteksi sheet sistem/hasil
  FULL_PROTECT.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.protect().setDescription('Sheet sistem/hasil - otomatis').setWarningOnly(true);
  });

  // 3. proteksi area tabel laporan (B2 tahun tetap bisa diedit)
  var lap = ss.getSheetByName(CFG.SHEET.LAPORAN);
  if (lap) lap.getRange('A4:G17').protect().setDescription('Tabel laba rugi - otomatis').setWarningOnly(true);

  // 4. proteksi area tampilan dashboard kecuali baris filter (row 7)
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
