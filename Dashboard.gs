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

  // lebar kolom & kanvas - lebih lega, total 14 kolom
  sh.setColumnWidth(1, 28);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 95);
  sh.setColumnWidth(14, 28);
  sh.getRange('A1:N60').setBackground(CFG.COLOR.BG_APP);

  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, PR=CFG.SHEET.PRODUK,
      UP=CFG.SHEET.UTANG_PIUTANG, AK=CFG.SHEET.AKUN, K=CFG.SHEET.KALKULASI;

  // ---------- HEADER (bar gradien + sub-header info) ----------
  sh.getRange('B2:M2').merge().setValue('💼  KEUANGAN KONTER PULSA')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontSize(20)
    .setFontFamily('Calibri').setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(2, 50);
  sh.getRange('B3:M3').merge().setFormula(
    '="📅  "&IF(D6="Semua","Tahun "&B6,TEXT(DATE(B6,IF(D6="Semua",1,D6),1),"mmmm yyyy"))' +
    '&"     |     📊  Mode: "&F6&"     |     💳  Akun: "&H6&"     |     🏷️  Kategori: "&J6&"     |     ✓  Status: "&L6')
    .setBackground('#1B3A5B').setFontColor('#E2E8F0').setFontSize(10)
    .setFontStyle('italic').setVerticalAlignment('middle');
  sh.setRowHeight(3, 24);
  sh.setRowHeight(4, 12);   // spacer

  // ---------- FILTER BAR ----------
  sh.getRange('B5:M5').merge().setValue('  ⚙️  FILTER  ')
    .setBackground('#E8EEF5').setFontColor(CFG.COLOR.HEADER_BG)
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('left');
  sh.setRowHeight(5, 22);
  var fLabels = [['Tahun',2],['Bulan',4],['Mode Waktu',6],['Akun',8],['Kategori',11],['Status',13]];
  fLabels.forEach(function(f){
    sh.getRange(6, f[1]).setValue(f[0]).setFontSize(8).setFontColor(CFG.COLOR.MUTED).setFontWeight('bold');
    sh.getRange(7, f[1]).setBackground('#FFFFFF').setFontWeight('bold').setFontColor(CFG.COLOR.ACCENT)
      .setHorizontalAlignment('center').setFontSize(11)
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.setRowHeight(6, 18);
  sh.setRowHeight(7, 26);

  // default value & validasi filter (geser ke baris 7)
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

  // pakai cell B7,D7,F7,H7,K7,M7 sebagai filter (perlu update _kalkulasi mirror)
  // (akan di-set di setupDashboardHelper_)

  // spacer
  sh.setRowHeight(8, 14);

  // ---------- KPI CARDS (4 + 4) ----------
  var dr = '">="&'+K+'!G1', dr2='"<="&'+K+'!G2';
  var omzet = '=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var labaKotor = '=SUMIFS('+JUAL+'!L2:L'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var pengeluaran = '=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  var adminFee = 'SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';

  // baris 9-11 = kartu atas, 12-14 = kartu bawah
  kpiCard_(sh, 9, 2,  '💰 OMZET',         'Total penjualan periode',     omzet,                       CFG.COLOR.ACCENT);
  kpiCard_(sh, 9, 5,  '📈 LABA KOTOR',    'Jual − HPP',                  labaKotor,                   CFG.COLOR.POS);
  kpiCard_(sh, 9, 8,  '💸 PENGELUARAN',   'Biaya operasional',           pengeluaran,                 CFG.COLOR.NEG);
  kpiCard_(sh, 9, 11, '🎯 LABA BERSIH',   'Kotor + Admin − Beban',       '=F10+'+adminFee+'-I10',     CFG.COLOR.POS);
  kpiCard_(sh, 12, 2, '🏦 SALDO KAS/BANK','Total saldo semua akun',      '=SUM('+AK+'!F2:F'+MR+')',   CFG.COLOR.HEADER_BG);
  kpiCard_(sh, 12, 5, '📦 NILAI STOK',    'Stok akhir × harga modal',    '=SUM('+PR+'!L2:L'+MR+')',   CFG.COLOR.HEADER_BG);
  kpiCard_(sh, 12, 8, '📥 PIUTANG',       'Tagihan ke pelanggan',        '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Piutang")', CFG.COLOR.WARN);
  kpiCard_(sh, 12,11, '📤 HUTANG',        'Kewajiban ke supplier',       '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Hutang")',  CFG.COLOR.WARN);

  sh.setRowHeight(15, 16);

  // ---------- SECTION: GRAFIK + ALERT ----------
  sh.getRange('B16:G16').merge().setValue('  📈  TREN OMZET & LABA  (ikut Mode Waktu)')
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11)
    .setBackground(CFG.COLOR.HEADER_BG).setHorizontalAlignment('left');
  sh.getRange('I16:M16').merge().setValue('  🔔  PANEL ALERT')
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11)
    .setBackground(CFG.COLOR.NEG).setHorizontalAlignment('left');
  sh.setRowHeight(16, 24);

  var kal = ss.getSheetByName(K);
  var chart = sh.newChart().asComboChart()
    .addRange(kal.getRange('N1:P54'))
    .setPosition(17, 2, 4, 4)
    .setOption('title', '')
    .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
    .setOption('width', 560).setOption('height', 320)
    .setOption('colors', [CFG.COLOR.ACCENT, CFG.COLOR.POS])
    .setOption('seriesType', 'bars')
    .setOption('series', { 1: { type: 'line', lineWidth: 3, pointSize: 5 } })
    .setOption('hAxis', { textStyle: { fontSize: 9 } })
    .setOption('vAxis', { format: 'short', textStyle: { fontSize: 9 } })
    .setOption('backgroundColor', CFG.COLOR.CARD)
    .setOption('chartArea', { left: 60, top: 20, width: '85%', height: '75%' })
    .build();
  sh.insertChart(chart);

  // PANEL ALERT (kolom I-M, baris 17 ke bawah)
  var alerts = [
    ['📦 Stok minus',         '=COUNTIF('+PR+'!N2:N'+MR+',"Minus")'],
    ['🟡 Stok habis',         '=COUNTIF('+PR+'!N2:N'+MR+',"Habis")'],
    ['🟠 Stok rendah',        '=COUNTIF('+PR+'!N2:N'+MR+',"Rendah")'],
    ['💸 Saldo akun negatif', '=COUNTIF('+AK+'!F2:F'+MR+',"<0")'],
    ['📥 Piutang jatuh tempo','=COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['📤 Hutang jatuh tempo', '=COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['⏳ Transaksi non-final','=COUNTIF('+JUAL+'!P2:P'+MR+',"Pending")+COUNTIF('+JUAL+'!P2:P'+MR+',"Draft")+COUNTIF('+JUAL+'!P2:P'+MR+',"Proses")'],
    ['⚠️ Penjualan warning',  '=COUNTIF('+JUAL+'!U2:U'+MR+',"?*")']
  ];
  for (var i = 0; i < alerts.length; i++) {
    var r = 17 + i;
    sh.getRange(r, 9, 1, 4).merge().setValue(alerts[i][0]).setFontSize(10)
      .setBackground('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 13).setFormula(alerts[i][1]).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(11)
      .setBackground('#FFFFFF')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r, 26);
  }
  // warna alert: >0 merah, 0 hijau
  var aRange = sh.getRange(17, 13, alerts.length, 1);
  var aRules = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([aRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0)
      .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([aRange]).build()
  ];
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat(aRules));

  sh.setRowHeight(35, 18);

  // ---------- SECTION: RINGKASAN BAWAH ----------
  var base = 36;
  sh.getRange(base, 2, 1, 11).merge().setValue('  📋  RINGKASAN PERIODE')
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11)
    .setBackground(CFG.COLOR.HEADER_BG).setHorizontalAlignment('left');
  sh.setRowHeight(base, 24);

  sectionTitle_(sh, base+1, 2, '🏦  SALDO PER AKUN');
  sh.getRange(base+2, 2).setFormula('=IFERROR(FILTER({'+AK+'!A2:A'+MR+','+AK+'!F2:F'+MR+'},'+AK+'!A2:A'+MR+'<>""),"")');
  sh.getRange(base+2, 3, 20, 1).setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0');

  sectionTitle_(sh, base+1, 5, '🏆  TOP PRODUK (Qty)');
  sh.getRange(base+2, 5).setFormula('=IFERROR(QUERY('+PR+'!A2:O'+MR+',"select B, I where I > 0 order by I desc limit 5",0),"Belum ada penjualan")');

  sectionTitle_(sh, base+1, 8, '💰  TOP LABA PRODUK');
  sh.getRange(base+2, 8).setFormula('=IFERROR(QUERY('+PR+'!A2:P'+MR+',"select B, P where P > 0 order by P desc limit 5",0),"Belum ada penjualan")');
  sh.getRange(base+2, 9, 6, 1).setNumberFormat('"Rp"#,##0');

  sectionTitle_(sh, base+1, 11, '💸  PENGELUARAN');
  sh.getRange(base+2, 11).setFormula('=IFERROR(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),"Belum ada pengeluaran")');
  sh.getRange(base+2, 12, 6, 1).setNumberFormat('"Rp"#,##0');

  // catatan audit
  sh.getRange(base+10, 2, 1, 11).merge().setFormula(
    '="📝  Catatan Audit: "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*CRITICAL*")&" kritis, "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*WARNING*")&" peringatan. Buka sheet audit untuk detail lengkap.   |   ⚙️ Klik menu Keuangan → 🔄 Refresh setelah input transaksi."')
    .setFontColor(CFG.COLOR.MUTED).setFontSize(9).setFontStyle('italic')
    .setBackground('#F1F5F9').setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(base+10, 22);
}

  // ---------------------------------------------------------------------------
//  Helper visual
// ---------------------------------------------------------------------------
function kpiCard_(sh, r, c, label, subtitle, formula, accent) {
  // baris r   : label (kecil)
  // baris r+1 : nilai (besar)
  // baris r+2 : subtitle (italic abu)
  sh.getRange(r, c, 1, 3).merge().setValue(label).setFontSize(9).setFontColor(CFG.COLOR.MUTED)
    .setFontWeight('bold').setBackground('#FFFFFF').setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sh.getRange(r+1, c, 1, 3).merge().setFormula(formula).setFontSize(16).setFontWeight('bold')
    .setFontColor(accent).setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r+2, c, 1, 3).merge().setValue(subtitle).setFontSize(8).setFontColor(CFG.COLOR.MUTED)
    .setFontStyle('italic').setBackground('#FFFFFF').setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  sh.getRange(r, c, 3, 3).setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(r, 18); sh.setRowHeight(r+1, 36); sh.setRowHeight(r+2, 18);
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

  // Mirror filter dashboard (filter ada di baris 7)
  k.getRange('D1').setValue('tahun');  k.getRange('E1').setFormula('='+D+'!B7');
  k.getRange('D2').setValue('bulan');  k.getRange('E2').setFormula('='+D+'!D7');
  k.getRange('D3').setValue('mode');   k.getRange('E3').setFormula('='+D+'!F7');
  k.getRange('D4').setValue('akun');   k.getRange('E4').setFormula('='+D+'!H7');
  k.getRange('D5').setValue('kategori');k.getRange('E5').setFormula('='+D+'!K7');
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
  // (Top list dipindah ke dashboard memakai QUERY - lebih andal)

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
