/**
 * ============================================================================
 *  Structure.gs  -  Membangun struktur 16 sheet (header, format, freeze)
 * ============================================================================
 */

function createAllSheets(ss) {
  // 1. Buat / pastikan semua sheet ada (urutan rapi)
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

  // 2. Hapus sheet bawaan "Sheet1" jika ada
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (def && order.indexOf(def.getName()) === -1) ss.deleteSheet(def);

  // 3. Tulis header + format untuk sheet bertabel
  Object.keys(HEADERS).forEach(function(key) {
    if (key === '_kalkulasi') return; // _kalkulasi ditata di Seed.gs
    writeHeader_(ss, key, HEADERS[key]);
  });

  // 4. Setup tampilan dasar tiap sheet data
  ['penjualan','pembelian','produk','transaksi_kas','koreksi_stok','utang_piutang',
   'buku_besar','akun','data_master','laporan','audit','skenario_uji','log_perubahan','_dropdown']
    .forEach(function(key){ styleDataSheet_(ss, key); });
}

/** Tulis baris header + format header (sheetKey == nama sheet) */
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

/** Format umum sheet data: border, lebar kolom otomatis, zebra ringan */
function styleDataSheet_(ss, sheetKey) {
  var sh = ss.getSheetByName(sheetKey);
  if (!sh) return;
  var nCol = HEADERS[sheetKey] ? HEADERS[sheetKey].length : sh.getLastColumn();
  if (!nCol) return;

  // lebar kolom
  for (var c = 1; c <= nCol; c++) {
    sh.setColumnWidth(c, 130);
  }
  // Hapus kolom & baris sisa agar bersih (sisakan 3 kolom helper sistem di kanan)
  var maxCols = sh.getMaxColumns();
  if (maxCols > nCol + 3) sh.deleteColumns(nCol + 4, maxCols - (nCol + 3));
  var maxRows = sh.getMaxRows();
  if (maxRows < CFG.MAX_ROW) sh.insertRowsAfter(maxRows, CFG.MAX_ROW - maxRows);

  // Format angka rupiah untuk kolom uang (dilakukan kasar; detail di Formulas.gs)
  sh.setHiddenGridlines ? null : null;
}

/** Terapkan format rupiah pada range tertentu */
function fmtRupiah_(sh, a1) {
  sh.getRange(a1).setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0');
}
function fmtTanggal_(sh, a1) {
  sh.getRange(a1).setNumberFormat('yyyy-mm-dd');
}
function fmtAngka_(sh, a1) {
  sh.getRange(a1).setNumberFormat('#,##0');
}
