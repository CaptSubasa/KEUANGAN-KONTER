/**
 * ============================================================================
 *  Dashboard.gs  -  Dashboard modern + 1 grafik dinamis (Mode Waktu) + laporan
 * ----------------------------------------------------------------------------
 *  Filter: Tahun, Bulan, Mode Waktu, Akun, Kategori, Status.
 *  Grafik tunggal membaca _kalkulasi (bucket berubah sesuai Mode Waktu).
 * ============================================================================
 */

function buildDashboard(ss) {
  setupDashboardHelper_(ss);   // siapkan _kalkulasi (filter, bucket, top list)

  var sh = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  sh.clear();
  sh.getCharts().forEach(function(c){ sh.removeChart(c); });
  sh.setHiddenGridlines(true);

  // lebar kolom & kanvas
  sh.setColumnWidth(1, 24);
  for (var c = 2; c <= 12; c++) sh.setColumnWidth(c, 96);
  sh.getRange('B1:L60').setBackground(CFG.COLOR.BG_APP);

  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, PR=CFG.SHEET.PRODUK,
      UP=CFG.SHEET.UTANG_PIUTANG, AK=CFG.SHEET.AKUN, K=CFG.SHEET.KALKULASI;

  // ---------- HEADER ----------
  sh.getRange('B2:L2').merge().setValue('KONTER PULSA  —  DASHBOARD KEUANGAN')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontSize(16)
    .setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(2, 40);
  sh.getRange('B3:L3').merge().setFormula(
    '="Periode: "&IF(D6="Semua","Tahun "&B6,TEXT(DATE(B6,IF(D6="Semua",1,D6),1),"mmmm")&" "&B6)' +
    '&"     •     Mode Grafik: "&F6&"     •     Akun: "&H6&"     •     Kategori: "&J6')
    .setBackground('#1B3A5B').setFontColor('#CBD5E1').setFontSize(10)
    .setVerticalAlignment('middle');
  sh.setRowHeight(3, 22);

  // ---------- FILTER ----------
  var fLabels = [['Tahun',2],['Bulan',4],['Mode Waktu',6],['Akun',8],['Kategori',10],['Status',12]];
  fLabels.forEach(function(f){
    sh.getRange(5, f[1]).setValue(f[0]).setFontSize(8).setFontColor(CFG.COLOR.MUTED).setFontWeight('bold');
    sh.getRange(6, f[1]).setBackground('#FFFFFF').setFontWeight('bold').setFontColor(CFG.COLOR.ACCENT)
      .setHorizontalAlignment('center')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  });

  // default value & validasi filter
  var now = new Date();
  sh.getRange('B6').setValue(now.getFullYear());
  sh.getRange('D6').setValue('Semua');
  sh.getRange('F6').setValue('Bulanan');
  sh.getRange('H6').setValue('Semua');
  sh.getRange('J6').setValue('Semua');
  sh.getRange('L6').setValue('Final');
  filterDV_(ss, sh, 'B6', K, 'AB');
  filterDV_(ss, sh, 'D6', K, 'AC');
  filterDV_(ss, sh, 'F6', CFG.SHEET.DROPDOWN, 'N');
  filterDV_(ss, sh, 'H6', K, 'AD');
  filterDV_(ss, sh, 'J6', K, 'AE');
  filterDV_(ss, sh, 'L6', K, 'AF');

  // ---------- KPI CARDS ----------
  var dr = '">="&'+K+'!G1', dr2='"<="&'+K+'!G2';
  var omzet = '=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var labaKotor = '=SUMIFS('+JUAL+'!L2:L'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var pengeluaran = '=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  var adminFee = 'SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';

  kpiCard_(sh, 8, 2,  'OMZET',          omzet,                       CFG.COLOR.ACCENT);
  kpiCard_(sh, 8, 5,  'LABA KOTOR',     labaKotor,                   CFG.COLOR.POS);
  kpiCard_(sh, 8, 8,  'PENGELUARAN',    pengeluaran,                 CFG.COLOR.NEG);
  kpiCard_(sh, 8, 11, 'LABA BERSIH',    '=E9+'+adminFee+'-H9',       CFG.COLOR.POS);
  kpiCard_(sh, 11, 2, 'SALDO KAS/BANK', '=SUMIF('+AK+'!A2:A'+MR+','+K+'!E9,'+AK+'!F2:F'+MR+')', CFG.COLOR.HEADER_BG);
  kpiCard_(sh, 11, 5, 'NILAI STOK',     '=SUMIF('+PR+'!C2:C'+MR+','+K+'!E8,'+PR+'!L2:L'+MR+')', CFG.COLOR.HEADER_BG);
  kpiCard_(sh, 11, 8, 'PIUTANG',        '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Piutang")', CFG.COLOR.WARN);
  kpiCard_(sh, 11, 11,'HUTANG',         '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Hutang")',  CFG.COLOR.WARN);

  // ---------- GRAFIK UTAMA (dinamis) ----------
  sh.getRange('B14:G14').merge().setValue('📈  TREN OMZET & LABA (mengikuti Mode Waktu)')
    .setFontWeight('bold').setFontColor(CFG.COLOR.HEADER_BG).setFontSize(11);
  var kal = ss.getSheetByName(K);
  var chart = sh.newChart().asColumnChart()
    .addRange(kal.getRange('N1:P54'))
    .setPosition(15, 2, 4, 4)
    .setOption('title', '')
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 560).setOption('height', 300)
    .setOption('colors', [CFG.COLOR.ACCENT, CFG.COLOR.POS])
    .setOption('hAxis', { textStyle: { fontSize: 9 } })
    .build();
  sh.insertChart(chart);

  // ---------- PANEL ALERT ----------
  sh.getRange('H14:L14').merge().setValue('🔔  PANEL ALERT')
    .setFontWeight('bold').setFontColor(CFG.COLOR.NEG).setFontSize(11);
  var alerts = [
    ['Stok minus',        '=COUNTIF('+PR+'!N2:N'+MR+',"Minus")'],
    ['Stok habis',        '=COUNTIF('+PR+'!N2:N'+MR+',"Habis")'],
    ['Stok rendah',       '=COUNTIF('+PR+'!N2:N'+MR+',"Rendah")'],
    ['Saldo akun negatif','=COUNTIF('+AK+'!F2:F'+MR+',"<0")'],
    ['Piutang jatuh tempo','=COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['Hutang jatuh tempo','=COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['Transaksi non-final','=COUNTIF('+JUAL+'!P2:P'+MR+',"Pending")+COUNTIF('+JUAL+'!P2:P'+MR+',"Draft")+COUNTIF('+JUAL+'!P2:P'+MR+',"Proses")'],
    ['Penjualan warning', '=COUNTIF('+JUAL+'!U2:U'+MR+',"?*")']
  ];
  for (var i = 0; i < alerts.length; i++) {
    var r = 15 + i;
    sh.getRange(r, 8, 1, 4).merge().setValue(alerts[i][0]).setFontSize(9)
      .setBackground('#FFFFFF').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 12).setFormula(alerts[i][1]).setHorizontalAlignment('center').setFontWeight('bold')
      .setBackground('#FFFFFF')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  }
  // warna alert: >0 merah, 0 hijau
  var aRange = sh.getRange(15, 12, alerts.length, 1);
  var aRules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([aRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0)
      .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([aRange]).build()
  ];
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat(aRules));

  // ---------- RINGKASAN (bawah) ----------
  var base = 33;
  sectionTitle_(sh, base, 2, '💼  SALDO PER AKUN');
  sh.getRange(base+1, 2).setFormula('=IFERROR(FILTER({'+AK+'!A2:A'+MR+','+AK+'!F2:F'+MR+'},'+AK+'!A2:A'+MR+'<>""),"")');
  sh.getRange(base+1, 3, 20, 1).setNumberFormat('"Rp"#,##0');

  sectionTitle_(sh, base, 5, '🏆  TOP PRODUK (Qty)');
  sh.getRange(base+1, 5).setFormula('='+K+'!R2:S6');

  sectionTitle_(sh, base, 8, '💰  TOP LABA PRODUK');
  sh.getRange(base+1, 8).setFormula('='+K+'!U2:V6');
  sh.getRange(base+1, 9, 6, 1).setNumberFormat('"Rp"#,##0');

  sectionTitle_(sh, base, 11, '💸  PENGELUARAN TERBESAR');
  sh.getRange(base+1, 11).setFormula('='+K+'!X2:Y6');
  sh.getRange(base+1, 12, 6, 1).setNumberFormat('"Rp"#,##0');

  // catatan audit
  sh.getRange(base+9, 2, 1, 11).merge().setFormula(
    '="Catatan Audit: "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*CRITICAL*")&" kritis, "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*WARNING*")&" peringatan. Buka sheet audit untuk detail."')
    .setFontColor(CFG.COLOR.MUTED).setFontSize(9).setFontStyle('italic');
}

// ---------------------------------------------------------------------------
//  Helper visual
// ---------------------------------------------------------------------------
function kpiCard_(sh, r, c, label, formula, accent) {
  sh.getRange(r, c, 1, 2).merge().setValue(label).setFontSize(8).setFontColor(CFG.COLOR.MUTED)
    .setFontWeight('bold').setBackground('#FFFFFF').setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sh.getRange(r+1, c, 1, 2).merge().setFormula(formula).setFontSize(14).setFontWeight('bold')
    .setFontColor(accent).setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r, c, 2, 2).setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 22); sh.setRowHeight(r+1, 34);
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

// ---------------------------------------------------------------------------
//  _kalkulasi : filter mirror, bucket grafik, top list, sumber filter
// ---------------------------------------------------------------------------
function setupDashboardHelper_(ss) {
  var k = ss.getSheetByName(CFG.SHEET.KALKULASI);
  if (k.getMaxColumns() < 40) k.insertColumnsAfter(k.getMaxColumns(), 40 - k.getMaxColumns());
  var D = CFG.SHEET.DASHBOARD, JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS,
      PR=CFG.SHEET.PRODUK, DM=CFG.SHEET.DATA_MASTER, AK=CFG.SHEET.AKUN;

  // Mirror filter dashboard
  k.getRange('D1').setValue('tahun');  k.getRange('E1').setFormula('='+D+'!B6');
  k.getRange('D2').setValue('bulan');  k.getRange('E2').setFormula('='+D+'!D6');
  k.getRange('D3').setValue('mode');   k.getRange('E3').setFormula('='+D+'!F6');
  k.getRange('D4').setValue('akun');   k.getRange('E4').setFormula('='+D+'!H6');
  k.getRange('D5').setValue('kategori');k.getRange('E5').setFormula('='+D+'!J6');
  k.getRange('E8').setFormula('=IF(E5="Semua","*",E5)');   // kat criteria
  k.getRange('E9').setFormula('=IF(E4="Semua","*",E4)');   // akun criteria

  // Batas tanggal KPI
  k.getRange('F1').setValue('startDate'); k.getRange('G1').setFormula('=IF(E2="Semua",DATE(E1,1,1),DATE(E1,E2,1))');
  k.getRange('F2').setValue('endDate');   k.getRange('G2').setFormula('=IF(E2="Semua",DATE(E1,12,31),EOMONTH(DATE(E1,E2,1),0))');

  // Bucket grafik (idx 1..53)
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

  // Omzet (O) & Laba (P) per bucket
  var fO = [], fP = [];
  for (var r = 2; r <= 54; r++) {
    fO.push(['=IF(L'+r+'="","",SUMIFS('+JUAL+'!$K$2:$K$'+MR+','+JUAL+'!$T$2:$T$'+MR+',"Final",'+JUAL+'!$A$2:$A$'+MR+',">="&L'+r+','+JUAL+'!$A$2:$A$'+MR+',"<="&M'+r+','+JUAL+'!$F$2:$F$'+MR+',$E$8))']);
    fP.push(['=IF(L'+r+'="","",SUMIFS('+JUAL+'!$L$2:$L$'+MR+','+JUAL+'!$T$2:$T$'+MR+',"Final",'+JUAL+'!$A$2:$A$'+MR+',">="&L'+r+','+JUAL+'!$A$2:$A$'+MR+',"<="&M'+r+','+JUAL+'!$F$2:$F$'+MR+',$E$8))']);
  }
  k.getRange(2, 15, 53, 1).setFormulas(fO);   // O2:O54
  k.getRange(2, 16, 53, 1).setFormulas(fP);   // P2:P54

  // Top list
  k.getRange('R2').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({'+PR+'!B2:B'+MR+',SUMIF('+JUAL+'!$E$2:$E$'+MR+','+PR+'!A2:A'+MR+','+JUAL+'!$V$2:$V$'+MR+')},'+PR+'!B2:B'+MR+'<>""),2,FALSE),5,2),"")');
  k.getRange('U2').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({'+PR+'!B2:B'+MR+',SUMIF('+JUAL+'!$E$2:$E$'+MR+','+PR+'!A2:A'+MR+','+JUAL+'!$W$2:$W$'+MR+')},'+PR+'!B2:B'+MR+'<>""),2,FALSE),5,2),"")');
  k.getRange('X2').setFormula('=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({'+DM+'!D2:D'+MR+',SUMIF('+TK+'!$D$2:$D$'+MR+','+DM+'!D2:D'+MR+','+TK+'!$Q$2:$Q$'+MR+')},'+DM+'!D2:D'+MR+'<>""),2,FALSE),5,2),"")');

  // Sumber filter dashboard
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

// ===========================================================================
//  LAPORAN LABA RUGI (bulanan, pilih tahun)
// ===========================================================================
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
  // Total
  var tr = 17;
  sh.getRange(tr, 1).setValue('TOTAL').setFontWeight('bold');
  for (var col = 2; col <= 7; col++) {
    sh.getRange(tr, col).setFormula('=SUM('+colLetter_(col)+'5:'+colLetter_(col)+'16)').setFontWeight('bold');
  }
  sh.getRange('B5:G17').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0');
  sh.getRange(tr, 1, 1, 7).setBackground('#E8EEF5');
  for (var w = 1; w <= 7; w++) sh.setColumnWidth(w, w === 1 ? 110 : 120);
}
