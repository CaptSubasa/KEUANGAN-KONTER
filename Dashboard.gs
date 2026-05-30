/**
 * ============================================================================
 *  Dashboard.gs  -  Dashboard modern + 1 grafik dinamis (Mode Waktu) + laporan
 * ----------------------------------------------------------------------------
 *  Filter: Tahun, Bulan, Mode Waktu, Akun, Kategori, Status.
 *  Grafik tunggal membaca _kalkulasi (bucket berubah sesuai Mode Waktu).
 * ============================================================================
 */

function buildDashboard(ss) {
  setupDashboardHelper_(ss);

  var sh = ss.getSheetByName(CFG.SHEET.DASHBOARD);
  sh.clear();
  sh.getCharts().forEach(function(c){ sh.removeChart(c); });
  sh.setHiddenGridlines(true);

  // === Kanvas: kolom A & N margin, B-M (12 kolom konten) ===
  sh.setColumnWidth(1, 22);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 96);
  sh.setColumnWidth(14, 22);
  sh.getRange('A1:N72').setBackground(CFG.COLOR.BG_APP);

  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, PR=CFG.SHEET.PRODUK,
      UP=CFG.SHEET.UTANG_PIUTANG, AK=CFG.SHEET.AKUN, K=CFG.SHEET.KALKULASI;

  // ===========================================================
  //  HEADER  (banner gelap + sub-info dengan timestamp)
  // ===========================================================
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

  // ===========================================================
  //  FILTER
  // ===========================================================
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

  // ===========================================================
  //  KPI TOP (4 cards, baris 9-12 : label/value/subtitle/sparkline)
  // ===========================================================
  var dr = '">="&'+K+'!G1', dr2='"<="&'+K+'!G2';
  var omzet = '=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var laba  = '=SUMIFS('+JUAL+'!L2:L'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var beban = '=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  var fee   = 'SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';

  // Top KPI: dengan sparkline trend
  premiumKPI_(sh, 9,  2, '💰  OMZET',         'Total penjualan periode', omzet,                   CFG.COLOR.ACCENT, K+'!O2:O54');
  premiumKPI_(sh, 9,  5, '📈  LABA KOTOR',    'Jual − HPP',              laba,                    CFG.COLOR.POS,    K+'!P2:P54');
  premiumKPI_(sh, 9,  8, '💸  PENGELUARAN',   'Biaya operasional',       beban,                   CFG.COLOR.NEG,    null);
  premiumKPI_(sh, 9, 11, '🎯  LABA BERSIH',   'Kotor + Admin − Beban',   '=B10+'+fee+'-H10',      CFG.COLOR.POS,    null);

  sh.setRowHeight(13, 14);

  // KPI BOTTOM (status keuangan)
  premiumKPI_(sh, 14,  2, '🏦  SALDO KAS/BANK', 'Total saldo akun',     '=SUM('+AK+'!F2:F'+MR+')',                                    CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 14,  5, '📦  NILAI STOK',     'Stok akhir × modal',   '=SUM('+PR+'!L2:L'+MR+')',                                    CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 14,  8, '📥  PIUTANG',        'Tagihan ke pelanggan', '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Piutang")',     CFG.COLOR.WARN,      null);
  premiumKPI_(sh, 14, 11, '📤  HUTANG',         'Kewajiban ke supplier','=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Hutang")',      CFG.COLOR.WARN,      null);

  sh.setRowHeight(18, 14);

  // ===========================================================
  //  SECTION: TREN OMZET & LABA  (chart utama, full width)
  // ===========================================================
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

  // ===========================================================
  //  SECTION: TOP PRODUK (kiri) + PANEL ALERT (kanan)
  // ===========================================================
  sh.getRange('B37:G37').merge().setValue('  🏆  TOP 5 PRODUK  ─  qty terjual')
    .setBackground(CFG.COLOR.POS).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.getRange('H37:M37').merge().setValue('  🔔  PANEL ALERT  ─  status sistem')
    .setBackground(CFG.COLOR.NEG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(37, 24);

  // --- TOP PRODUK (kolom B-G, header + 5 baris) ---
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

  // --- PANEL ALERT (kolom H-M) ---
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
  // Conditional formatting alert
  var aRange = sh.getRange(38, 13, alerts.length, 1);
  var existing = sh.getConditionalFormatRules();
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setBackground('#FEE2E2').setFontColor(CFG.COLOR.NEG).setRanges([aRange]).build());
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0)
    .setBackground('#DCFCE7').setFontColor(CFG.COLOR.POS).setRanges([aRange]).build());
  sh.setConditionalFormatRules(existing);

  sh.setRowHeight(45, 14);

  // ===========================================================
  //  SECTION: SALDO PER AKUN + PENGELUARAN TERBESAR
  // ===========================================================
  sh.getRange('B46:G46').merge().setValue('  🏦  SALDO PER AKUN  ─  dengan persentase')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.getRange('H46:M46').merge().setValue('  💸  TOP 5 PENGELUARAN  ─  per kategori')
    .setBackground(CFG.COLOR.NEG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(46, 24);

  // --- SALDO PER AKUN ---
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

  // --- PENGELUARAN TERBESAR ---
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

  // ===========================================================
  //  FOOTER
  // ===========================================================
  sh.setRowHeight(56, 12);
  sh.getRange('B57:M57').merge().setFormula(
    '="📝  Catatan Audit: "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*CRITICAL*")&" kritis, "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*WARNING*")&" peringatan.   ●   ⚙️ Klik menu Keuangan → 🔄 Refresh setelah input transaksi.   ●   Filter di atas mengubah seluruh angka dashboard."')
    .setFontColor('#475569').setFontSize(9).setFontStyle('italic')
    .setBackground('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(57, 28);
}

// ---------------------------------------------------------------------------
//  Helper: Premium KPI card  (label / nilai besar / subtitle / sparkline)
// ---------------------------------------------------------------------------
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
