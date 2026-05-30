/**
 * ============================================================================
 *  Formulas.gs  -  Formula inti (ARRAYFORMULA s/d baris 5000) + audit
 *  Prinsip: kolom formula tidak menyentuh status non-final.
 *  Status final ditentukan terpusat lewat named range FINAL_STATUS.
 * ============================================================================
 */

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

// ---------------------------------------------------------------------------
//  PENJUALAN
// ---------------------------------------------------------------------------
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

  // helper sistem (kolom V) : qty efektif keluar (hanya status final)
  sh.getRange('V1').setValue('eff_keluar (sys)');
  setAF_(sh, 'V2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(T2:T'+MR+'="Final",N(G2:G'+MR+'),0)))');
  // helper sistem (kolom W) : laba efektif (hanya status final)
  sh.getRange('W1').setValue('eff_laba (sys)');
  setAF_(sh, 'W2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(T2:T'+MR+'="Final",L2:L'+MR+',0)))');
  sh.hideColumns(22, 2);
}

// ---------------------------------------------------------------------------
//  PEMBELIAN
// ---------------------------------------------------------------------------
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

  // helper sistem (kolom R) : qty efektif masuk (hanya status final)
  sh.getRange('R1').setValue('eff_masuk (sys)');
  setAF_(sh, 'R2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(COUNTIF(FINAL_STATUS,L2:L'+MR+')>0,N(G2:G'+MR+'),0)))');
  sh.hideColumns(18);
}

// ---------------------------------------------------------------------------
//  PRODUK (stok berjalan)
// ---------------------------------------------------------------------------
function formulasProduk_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.PRODUK);
  var JUAL = CFG.SHEET.PENJUALAN, BELI = CFG.SHEET.PEMBELIAN, KOR = CFG.SHEET.KOREKSI_STOK;

  // Stok Masuk (H) <- pembelian.eff_masuk (R) per kode (A)
  setAF_(sh, 'H2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+BELI+'!$E$2:$E$'+MR+',A2:A'+MR+','+BELI+'!$R$2:$R$'+MR+')))');
  // Stok Keluar (I) <- penjualan.eff_keluar (V) per kode (E)
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+JUAL+'!$E$2:$E$'+MR+',A2:A'+MR+','+JUAL+'!$V$2:$V$'+MR+')))');
  // Koreksi (J) <- koreksi_stok.eff (K) per kode (D)
  setAF_(sh, 'J2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",SUMIF('+KOR+'!$D$2:$D$'+MR+',A2:A'+MR+','+KOR+'!$K$2:$K$'+MR+')))');
  // Stok Akhir (K) = Awal + Masuk - Keluar + Koreksi
  setAF_(sh, 'K2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",N(G2:G'+MR+')+H2:H'+MR+'-I2:I'+MR+'+J2:J'+MR+'))');
  // Nilai Stok (L) = Stok Akhir x Harga Modal Default (E)
  setAF_(sh, 'L2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",K2:K'+MR+'*N(E2:E'+MR+')))');
  // Status Stok (N)
  setAF_(sh, 'N2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(K2:K'+MR+'<0,"Minus",IF(K2:K'+MR+'=0,"Habis",IF(K2:K'+MR+'<=N(M2:M'+MR+'),"Rendah","Aman")))))');
  // Warning (O)
  setAF_(sh, 'O2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(K2:K'+MR+'<0,"⚠️ Stok minus",IF(K2:K'+MR+'<=N(M2:M'+MR+'),"⚠️ Stok menipis",""))))');
}

// ---------------------------------------------------------------------------
//  TRANSAKSI_KAS
// ---------------------------------------------------------------------------
function formulasTransaksiKas_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS);

  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","TKS-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  // Laba Admin (I) = admin fee hanya untuk tarik_tunai / transfer_uang DAN status final
  setAF_(sh, 'I2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(((C2:C'+MR+'="tarik_tunai")+(C2:C'+MR+'="transfer_uang")>0)*(COUNTIF(FINAL_STATUS,K2:K'+MR+')>0),N(H2:H'+MR+'),0)))');
  // Warning (P)
  setAF_(sh, 'P2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF((C2:C'+MR+'="koreksi_kas")*(COUNTIF(FINAL_STATUS,K2:K'+MR+')=0),"⚠️ Koreksi kas belum Approved",'+
    'IF((C2:C'+MR+'="pengeluaran")*(E2:E'+MR+'=""),"⚠️ Akun Keluar kosong",'+
    'IF((C2:C'+MR+'="modal")*(F2:F'+MR+'=""),"⚠️ Akun Masuk kosong",'+
    'IF(((C2:C'+MR+'="mutasi")+(C2:C'+MR+'="tarik_tunai")+(C2:C'+MR+'="transfer_uang")>0)*((E2:E'+MR+'="")+(F2:F'+MR+'="")>0),"⚠️ Akun dari/ke kosong",'+
    'IF(N(G2:G'+MR+')<=0,"⚠️ Nominal kosong/0","")))))))');

  // helper sistem (kolom Q) : pengeluaran efektif (jenis=pengeluaran & status final)
  sh.getRange('Q1').setValue('eff_pengeluaran (sys)');
  setAF_(sh, 'Q2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF((C2:C'+MR+'="pengeluaran")*(COUNTIF(FINAL_STATUS,K2:K'+MR+')>0),N(G2:G'+MR+'),0)))');
  sh.hideColumns(17);
}

// ---------------------------------------------------------------------------
//  KOREKSI_STOK
// ---------------------------------------------------------------------------
function formulasKoreksiStok_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.KOREKSI_STOK);
  var P = CFG.SHEET.PRODUK;

  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","KST-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  setAF_(sh, 'D2', '=ARRAYFORMULA(IF(C2:C'+MR+'="","",IFERROR(INDEX('+P+'!$A$2:$A$'+MR+',MATCH(C2:C'+MR+','+P+'!$B$2:$B$'+MR+',0)),"❓ tdk ada")))');
  // Warning (J)
  setAF_(sh, 'J2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(C2:C'+MR+'="","⚠️ Produk kosong",'+
    'IF(N(F2:F'+MR+')<=0,"⚠️ Qty kosong/0",'+
    'IF(COUNTIF(FINAL_STATUS,H2:H'+MR+')=0,"⏳ Menunggu Approve","✅ Diterapkan")))))');

  // helper sistem (kolom K) : koreksi efektif (hanya Approved); Tambah=+, Kurang=-
  sh.getRange('K1').setValue('eff_koreksi (sys)');
  setAF_(sh, 'K2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",IF(COUNTIF(FINAL_STATUS,H2:H'+MR+')>0,'+
    'IF(E2:E'+MR+'="Tambah",N(F2:F'+MR+'),-N(F2:F'+MR+')),0)))');
  sh.hideColumns(11);
}

// ---------------------------------------------------------------------------
//  UTANG_PIUTANG
// ---------------------------------------------------------------------------
function formulasUtangPiutang_(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG);

  setAF_(sh, 'B2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","","UPI-"&TEXT(A2:A'+MR+',"YYYYMMDD")&"-"&TEXT(ROW(A2:A'+MR+')-1,"0000")))');
  // Sisa (H) = Nominal - Sudah Dibayar
  setAF_(sh, 'H2', '=ARRAYFORMULA(IF(A2:A'+MR+'="","",N(E2:E'+MR+')-N(G2:G'+MR+')))');
  // Status Jatuh Tempo (K)
  setAF_(sh, 'K2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF(H2:H'+MR+'<=0,"✅ Lunas",'+
    'IF(F2:F'+MR+'="","-",'+
    'IF(F2:F'+MR+'<TODAY(),"❗ Telat",'+
    'IF(F2:F'+MR+'<=TODAY()+3,"⚠️ Jatuh Tempo","Aman"))))))');
  // Warning (M)
  setAF_(sh, 'M2',
    '=ARRAYFORMULA(IF(A2:A'+MR+'="","",'+
    'IF((H2:H'+MR+'>0)*(F2:F'+MR+'<>"")*(F2:F'+MR+'<TODAY()),"❗ Lewat jatuh tempo","")))');
}

// ---------------------------------------------------------------------------
//  AKUN (saldo dari buku_besar)
// ---------------------------------------------------------------------------
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

// ---------------------------------------------------------------------------
//  FORMAT ANGKA
// ---------------------------------------------------------------------------
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

// ===========================================================================
//  AUDIT  -  15 pemeriksaan otomatis
// ===========================================================================
function buildAuditDefinitions(ss) {
  var sh = ss.getSheetByName(CFG.SHEET.AUDIT);
  var J = CFG.SHEET.PENJUALAN, B = CFG.SHEET.PEMBELIAN, P = CFG.SHEET.PRODUK,
      T = CFG.SHEET.TRANSAKSI_KAS, K = CFG.SHEET.KOREKSI_STOK, U = CFG.SHEET.UTANG_PIUTANG,
      A = CFG.SHEET.AKUN, BB = CFG.SHEET.BUKU_BESAR;

  // [pemeriksaan, formulaJumlah, severity, detail]
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

  // Pewarnaan status
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
