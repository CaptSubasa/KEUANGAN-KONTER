/**
 * ============================================================================
 *  Validation.gs  -  Dropdown s/d baris 5000 (sumber dari _dropdown)
 *  Tidak ada hardcode list di Data Validation; semua dari range _dropdown.
 *  setAllowInvalid(true) -> input tidak diblokir keras (cukup peringatan).
 * ============================================================================
 */

function applyValidation(ss) {
  fillDropdownSource_(ss);

  var DD = CFG.SHEET.DROPDOWN;

  // ---- penjualan ----
  dv_(ss, CFG.SHEET.PENJUALAN, 'C', DD, 'I');   // Jenis Transaksi
  dv_(ss, CFG.SHEET.PENJUALAN, 'D', DD, 'A');   // Nama Produk
  dv_(ss, CFG.SHEET.PENJUALAN, 'M', DD, 'D');   // Metode
  dv_(ss, CFG.SHEET.PENJUALAN, 'N', DD, 'B');   // Akun Uang Masuk
  dv_(ss, CFG.SHEET.PENJUALAN, 'O', DD, 'B');   // Akun Modal / Sumber HPP
  dv_(ss, CFG.SHEET.PENJUALAN, 'P', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.PENJUALAN, 'Q', DD, 'F');   // Pelanggan
  dv_(ss, CFG.SHEET.PENJUALAN, 'S', DD, 'G');   // User

  // ---- pembelian ----
  dv_(ss, CFG.SHEET.PEMBELIAN, 'C', DD, 'E');   // Supplier
  dv_(ss, CFG.SHEET.PEMBELIAN, 'D', DD, 'A');   // Produk
  dv_(ss, CFG.SHEET.PEMBELIAN, 'J', DD, 'D');   // Metode
  dv_(ss, CFG.SHEET.PEMBELIAN, 'K', DD, 'B');   // Akun Pembayaran
  dv_(ss, CFG.SHEET.PEMBELIAN, 'L', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.PEMBELIAN, 'M', DD, 'M');   // Jenis Pembelian
  dv_(ss, CFG.SHEET.PEMBELIAN, 'P', DD, 'G');   // User

  // ---- produk ----
  dv_(ss, CFG.SHEET.PRODUK, 'C', DD, 'H');      // Kategori

  // ---- transaksi_kas ----
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'C', DD, 'J');  // Jenis
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'D', DD, 'P');  // Kategori
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'E', DD, 'B');  // Akun Keluar
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'F', DD, 'B');  // Akun Masuk
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'J', DD, 'D');  // Metode
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'K', DD, 'C');  // Status
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'L', DD, 'O');  // Arah Koreksi
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'M', DD, 'F');  // Pihak (saran pelanggan)
  dv_(ss, CFG.SHEET.TRANSAKSI_KAS, 'O', DD, 'G');  // User

  // ---- koreksi_stok ----
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'C', DD, 'A');   // Produk
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'E', DD, 'K');   // Arah (Tambah/Kurang)
  dvDirect_(ss, CFG.SHEET.KOREKSI_STOK, 'G', CFG.SHEET.DATA_MASTER, 'I'); // Alasan = Jenis Koreksi Stok
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'H', DD, 'C');   // Status
  dv_(ss, CFG.SHEET.KOREKSI_STOK, 'I', DD, 'G');   // User

  // ---- utang_piutang ----
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'C', DD, 'L');  // Arah (Piutang/Hutang)
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'D', DD, 'F');  // Pihak
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'I', DD, 'B');  // Akun Bayar
  dv_(ss, CFG.SHEET.UTANG_PIUTANG, 'J', DD, 'C');  // Status

  // ---- akun ----
  dv_(ss, CFG.SHEET.AKUN, 'B', DD, 'Q');           // Jenis Akun
}

// ---------------------------------------------------------------------------
//  Isi sumber dropdown (_dropdown)
// ---------------------------------------------------------------------------
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

  // Enum sistem (literal, tetap berbasis range)
  setCol_(dd, 11, ['Tambah','Kurang']);                          // K Arah Stok
  setCol_(dd, 12, ['Piutang','Hutang']);                         // L Arah Utang/Piutang
  setCol_(dd, 13, ['Tunai','Kredit']);                           // M Jenis Pembelian
  setCol_(dd, 14, ['Harian','Mingguan','Bulanan','Tahunan']);    // N Mode Waktu
  setCol_(dd, 15, ['Tambah','Kurang']);                          // O Arah Koreksi Kas
  // P Kategori Kas = gabungan kategori pengeluaran + jenis mutasi + jenis koreksi kas
  dd.getRange('P2').setFormula(
    '=IFERROR(FILTER({'+DM+'!D2:D'+M+';'+DM+'!G2:G'+M+';'+DM+'!H2:H'+M+'},'+
    '{'+DM+'!D2:D'+M+';'+DM+'!G2:G'+M+';'+DM+'!H2:H'+M+'}<>""),"")');
  setCol_(dd, 17, ['Kas','Bank','E-Wallet','QRIS','Supplier','Lainnya']);  // Q Jenis Akun
}

// ---------------------------------------------------------------------------
//  Helper penerap validasi
// ---------------------------------------------------------------------------
function dv_(ss, targetSheet, targetCol, srcSheet, srcCol) {
  var src = ss.getSheetByName(srcSheet).getRange(srcCol + '2:' + srcCol + '1000');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(src, true).setAllowInvalid(true).build();
  ss.getSheetByName(targetSheet).getRange(targetCol + '2:' + targetCol + MR).setDataValidation(rule);
}

function dvDirect_(ss, targetSheet, targetCol, srcSheet, srcCol) {
  dv_(ss, targetSheet, targetCol, srcSheet, srcCol);
}
