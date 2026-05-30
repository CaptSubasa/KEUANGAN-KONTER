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

  // === Kanvas: 14 kolom (A & N margin, B-M konten = 12 kolom) ===
  sh.setColumnWidth(1, 24);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 102);
  sh.setColumnWidth(14, 24);
  sh.getRange('A1:N75').setBackground(CFG.COLOR.BG_APP);

  var JUAL=CFG.SHEET.PENJUALAN, TK=CFG.SHEET.TRANSAKSI_KAS, PR=CFG.SHEET.PRODUK,
      UP=CFG.SHEET.UTANG_PIUTANG, AK=CFG.SHEET.AKUN, K=CFG.SHEET.KALKULASI;

  // ===========================================================
  //  HEADER  (banner gelap, judul besar, sub-info elegant)
  // ===========================================================
  sh.getRange('B2:M2').merge().setValue('  KEUANGAN KONTER PULSA  •  DASHBOARD')
    .setBackground(CFG.COLOR.HEADER_BG).setFontColor('#FFFFFF').setFontSize(20).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setFontFamily('Calibri');
  sh.setRowHeight(2, 60);

  sh.getRange('B3:M3').merge().setFormula(
    '="  📅  "&IF(D7="Semua","Tahun "&B7,TEXT(DATE(B7,IF(D7="Semua",1,D7),1),"mmmm yyyy"))' +
    '&"   •   "&F7&"   •   "&H7&"   •   "&K7&"   •   ✓ "&M7' +
    '&"   •   Update: "&TEXT(NOW(),"dd-mmm hh:mm")')
    .setBackground('#1E293B').setFontColor('#94A3B8').setFontSize(10)
    .setVerticalAlignment('middle').setFontFamily('Calibri');
  sh.setRowHeight(3, 26);
  sh.setRowHeight(4, 14);

  // ===========================================================
  //  FILTER (label kapital + dropdown card)
  // ===========================================================
  sh.getRange('B5:M5').merge().setValue('  FILTER')
    .setBackground('#FFFFFF').setFontColor(CFG.COLOR.MUTED).setFontWeight('bold').setFontSize(9)
    .setBorder(false, false, true, false, false, false, CFG.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(5, 22);

  var fLabels = [['TAHUN',2],['BULAN',4],['MODE WAKTU',6],['AKUN',8],['KATEGORI',11],['STATUS',13]];
  fLabels.forEach(function(f){
    sh.getRange(6, f[1]).setValue(f[0]).setFontSize(8).setFontColor(CFG.COLOR.SUBMUTED)
      .setFontWeight('bold').setBackground('#FFFFFF');
    sh.getRange(7, f[1]).setBackground('#FFFFFF').setFontWeight('bold').setFontColor(CFG.COLOR.TEXT_DARK)
      .setHorizontalAlignment('center').setFontSize(11).setFontFamily('Calibri')
      .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange('B5:M7').setBackground('#FFFFFF');
  sh.getRange('B7:M7').setBorder(false,false,true,false,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(6, 18); sh.setRowHeight(7, 32); sh.setRowHeight(8, 18);

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
  //  KPI TOP  (4 cards dengan accent stripe + sparkline) row 9-13
  // ===========================================================
  var dr = '">="&'+K+'!G1', dr2='"<="&'+K+'!G2';
  var omzet = '=SUMIFS('+JUAL+'!K2:K'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var laba  = '=SUMIFS('+JUAL+'!L2:L'+MR+','+JUAL+'!T2:T'+MR+',"Final",'+JUAL+'!A2:A'+MR+','+dr+','+JUAL+'!A2:A'+MR+','+dr2+','+JUAL+'!F2:F'+MR+','+K+'!E8)';
  var beban = '=SUMIFS('+TK+'!Q2:Q'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';
  var fee   = 'SUMIFS('+TK+'!I2:I'+MR+','+TK+'!A2:A'+MR+','+dr+','+TK+'!A2:A'+MR+','+dr2+')';

  premiumKPI_(sh, 9,  2, 'OMZET',           'Total penjualan periode', omzet,                   CFG.COLOR.ACCENT, K+'!O2:O54');
  premiumKPI_(sh, 9,  5, 'LABA KOTOR',      'Pendapatan bersih jual',  laba,                    CFG.COLOR.POS,    K+'!P2:P54');
  premiumKPI_(sh, 9,  8, 'PENGELUARAN',     'Biaya operasional',       beban,                   CFG.COLOR.NEG,    null);
  premiumKPI_(sh, 9, 11, 'LABA BERSIH',     'Kotor + Admin − Beban',   '=B11+'+fee+'-H11',      CFG.COLOR.POS,    null);

  sh.setRowHeight(14, 14);

  // KPI BOTTOM (cards dengan accent stripe, no sparkline) row 15-19
  premiumKPI_(sh, 15,  2, 'SALDO KAS / BANK', 'Total saldo akun',     '=SUM('+AK+'!F2:F'+MR+')',                                   CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 15,  5, 'NILAI STOK',       'Stok akhir × modal',   '=SUM('+PR+'!L2:L'+MR+')',                                   CFG.COLOR.HEADER_BG, null);
  premiumKPI_(sh, 15,  8, 'PIUTANG',          'Tagihan pelanggan',    '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Piutang")',    CFG.COLOR.WARN,      null);
  premiumKPI_(sh, 15, 11, 'HUTANG',           'Kewajiban supplier',   '=SUMIFS('+UP+'!H2:H'+MR+','+UP+'!C2:C'+MR+',"Hutang")',     CFG.COLOR.WARN,      null);

  sh.setRowHeight(20, 16);

  // ===========================================================
  //  SECTION: TREN  (chart full width)
  // ===========================================================
  sectionHeader_(sh, 'B21:M21', 'TREN OMZET & LABA', 'grafik mengikuti Mode Waktu', CFG.COLOR.ACCENT);
  sh.setRowHeight(21, 30);

  var kal = ss.getSheetByName(K);
  var chart = sh.newChart().asComboChart()
    .addRange(kal.getRange('N1:P54'))
    .setPosition(22, 2, 6, 6)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'end',
                           textStyle: { fontSize: 11, color: '#475569', bold: false } })
    .setOption('width', 1180).setOption('height', 340)
    .setOption('colors', [CFG.COLOR.ACCENT, CFG.COLOR.POS])
    .setOption('seriesType', 'bars')
    .setOption('series', {
      0: { labelInLegend: 'Omzet', color: CFG.COLOR.ACCENT },
      1: { type: 'line', lineWidth: 3, pointSize: 6, labelInLegend: 'Laba Kotor', color: CFG.COLOR.POS,
           pointShape: 'circle' }
    })
    .setOption('hAxis', { textStyle: { fontSize: 10, color: '#64748B' }, baselineColor: '#CBD5E1' })
    .setOption('vAxis', { format: 'short', textStyle: { fontSize: 10, color: '#64748B' },
                          gridlines: { color: '#F1F5F9', count: 5 }, baselineColor: '#CBD5E1' })
    .setOption('backgroundColor', { fill: '#FFFFFF', stroke: '#E2E8F0', strokeWidth: 0 })
    .setOption('chartArea', { left: 70, top: 36, width: '88%', height: '76%' })
    .setOption('bar', { groupWidth: '55%' })
    .setOption('focusTarget', 'category')
    .build();
  sh.insertChart(chart);
  for (var rr = 22; rr <= 36; rr++) sh.setRowHeight(rr, 22);
  sh.setRowHeight(37, 16);

  // ===========================================================
  //  TOP PRODUK + PANEL ALERT (2 kolom)
  // ===========================================================
  sectionHeader_(sh, 'B38:G38', 'TOP 5 PRODUK', 'qty terjual periode', CFG.COLOR.POS);
  sectionHeader_(sh, 'H38:M38', 'PANEL ALERT', 'status sistem terkini', CFG.COLOR.NEG);
  sh.setRowHeight(38, 30);

  // --- TOP PRODUK ---
  tableHeader_(sh, 39, 2, 3, 'Produk', 'left');
  tableHeader_(sh, 39, 5, 1, 'Qty', 'center');
  tableHeader_(sh, 39, 6, 2, 'Visualisasi', 'left');
  sh.setRowHeight(39, 24);

  for (var i = 0; i < 5; i++) {
    var r = 40 + i;
    sh.getRange(r, 2, 1, 3).merge().setFormula(
      '=IFERROR(IF(INDEX(QUERY('+PR+'!A2:O'+MR+',"select B where I > 0 order by I desc limit 5",0),'+(i+1)+',1)="","—",INDEX(QUERY('+PR+'!A2:O'+MR+',"select B where I > 0 order by I desc limit 5",0),'+(i+1)+',1)),"—")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle')
      .setFontColor(CFG.COLOR.TEXT_DARK).setHorizontalAlignment('left');
    sh.getRange(r, 5).setFormula(
      '=IFERROR(INDEX(QUERY('+PR+'!A2:O'+MR+',"select I where I > 0 order by I desc limit 5",0),'+(i+1)+',1),"")')
      .setBackground('#FFFFFF').setHorizontalAlignment('center').setNumberFormat('#,##0')
      .setFontWeight('bold').setFontColor(CFG.COLOR.ACCENT).setFontSize(10);
    sh.getRange(r, 6, 1, 2).merge().setFormula(
      '=IFERROR(IF(MAX($E$40:$E$44)=0,"",REPT("▮",ROUND(E'+r+'/MAX($E$40:$E$44)*12,0))),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.POS).setFontSize(11)
      .setVerticalAlignment('middle');
    sh.setRowHeight(r, 24);
    // alternate row bg
    if (i % 2 === 1) sh.getRange(r, 2, 1, 6).setBackground('#FAFBFC');
  }
  sh.getRange('B39:G44').setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);

  // --- PANEL ALERT ---
  var alerts = [
    ['Stok minus',           '=COUNTIF('+PR+'!N2:N'+MR+',"Minus")'],
    ['Stok habis',           '=COUNTIF('+PR+'!N2:N'+MR+',"Habis")'],
    ['Stok rendah',          '=COUNTIF('+PR+'!N2:N'+MR+',"Rendah")'],
    ['Saldo akun negatif',   '=COUNTIF('+AK+'!F2:F'+MR+',"<0")'],
    ['Piutang jatuh tempo',  '=COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Piutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['Hutang jatuh tempo',   '=COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Telat*")+COUNTIFS('+UP+'!C2:C'+MR+',"Hutang",'+UP+'!K2:K'+MR+',"*Jatuh Tempo*")'],
    ['Transaksi non-final',  '=COUNTIF('+JUAL+'!P2:P'+MR+',"Pending")+COUNTIF('+JUAL+'!P2:P'+MR+',"Draft")+COUNTIF('+JUAL+'!P2:P'+MR+',"Proses")']
  ];
  // alert dimulai di row 39 dengan 1 row spacer (39 = title), sejajar table top produk
  // total 6 baris alert (39 sampai 44 atau lebih) - kita ratakan dengan top produk: 5 baris (40-44)
  // tampilkan 6 alert + 1 ringkasan total -> ada 7 baris, lebih pendek dari section, ok
  // Use rows 39-45 for alert (7 rows, 6 alerts + total)
  for (var i = 0; i < alerts.length; i++) {
    var r = 39 + i;
    sh.getRange(r, 8, 1, 4).merge().setValue(alerts[i][0]).setFontSize(10)
      .setBackground('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setFontColor(CFG.COLOR.TEXT_DARK)
      .setBorder(true,true,true,false,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    // status badge style
    sh.getRange(r, 12).setFormula('=IF('+alerts[i][1].substring(1)+'>0,"⚠ "&'+alerts[i][1].substring(1)+',"✓")')
      .setHorizontalAlignment('center').setFontWeight('bold').setFontSize(10)
      .setBackground('#FFFFFF').setFontFamily('Calibri')
      .setBorder(true,false,true,false,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 13).setFormula(alerts[i][1]).setHorizontalAlignment('center').setFontSize(10)
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.SUBMUTED)
      .setBorder(true,false,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r, 22);
    if (i % 2 === 1) sh.getRange(r, 8, 1, 6).setBackground('#FAFBFC');
  }
  // conditional formatting badge L column (col 12)
  var badgeRange = sh.getRange(39, 12, alerts.length, 1);
  var existing = sh.getConditionalFormatRules();
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('⚠')
    .setBackground(CFG.COLOR.NEG_LT).setFontColor(CFG.COLOR.NEG).setRanges([badgeRange]).build());
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('✓')
    .setBackground(CFG.COLOR.POS_LT).setFontColor(CFG.COLOR.POS).setRanges([badgeRange]).build());
  sh.setConditionalFormatRules(existing);

  sh.setRowHeight(46, 16);

  // ===========================================================
  //  SALDO PER AKUN + PENGELUARAN TERBESAR
  // ===========================================================
  sectionHeader_(sh, 'B47:G47', 'SALDO PER AKUN', 'distribusi saldo', CFG.COLOR.HEADER_BG);
  sectionHeader_(sh, 'H47:M47', 'TOP 5 PENGELUARAN', 'per kategori', CFG.COLOR.NEG);
  sh.setRowHeight(47, 30);

  // SALDO PER AKUN
  tableHeader_(sh, 48, 2, 2, 'Akun', 'left');
  tableHeader_(sh, 48, 4, 1, 'Jenis', 'left');
  tableHeader_(sh, 48, 5, 1, 'Saldo', 'right');
  tableHeader_(sh, 48, 6, 2, 'Distribusi', 'left');
  sh.setRowHeight(48, 24);

  for (var i = 0; i < 8; i++) {
    var r = 49 + i;
    sh.getRange(r, 2, 1, 2).merge().setFormula('=IFERROR(IF(INDEX('+AK+'!A2:A,'+(i+1)+')="","",INDEX('+AK+'!A2:A,'+(i+1)+')),"")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle')
      .setFontColor(CFG.COLOR.TEXT_DARK).setFontWeight('bold');
    sh.getRange(r, 4).setFormula('=IFERROR(INDEX('+AK+'!B2:B,'+(i+1)+'),"")')
      .setBackground('#FFFFFF').setFontSize(9).setFontColor(CFG.COLOR.SUBMUTED).setVerticalAlignment('middle');
    sh.getRange(r, 5).setFormula('=IFERROR(IF(INDEX('+AK+'!A2:A,'+(i+1)+')="","",INDEX('+AK+'!F2:F,'+(i+1)+')),"")')
      .setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0').setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontSize(10)
      .setFontColor(CFG.COLOR.TEXT_DARK);
    sh.getRange(r, 6, 1, 2).merge().setFormula(
      '=IFERROR(IF(OR(E'+r+'="",E'+r+'<=0,SUM('+AK+'!F2:F'+MR+')<=0),"",REPT("▮",ROUND(E'+r+'/SUM('+AK+'!F2:F'+MR+')*12,0))&"  "&TEXT(E'+r+'/SUM('+AK+'!F2:F'+MR+'),"0%")),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.ACCENT).setFontSize(10)
      .setVerticalAlignment('middle');
    sh.setRowHeight(r, 22);
    if (i % 2 === 1) sh.getRange(r, 2, 1, 6).setBackground('#FAFBFC');
  }
  sh.getRange('B48:G56').setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);

  // PENGELUARAN TERBESAR
  tableHeader_(sh, 48, 8, 2, 'Kategori', 'left');
  tableHeader_(sh, 48, 10, 1, 'Nominal', 'right');
  tableHeader_(sh, 48, 11, 3, 'Visualisasi', 'left');

  for (var i = 0; i < 5; i++) {
    var r = 49 + i;
    sh.getRange(r, 8, 1, 2).merge().setFormula(
      '=IFERROR(IF(INDEX(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),'+(i+1)+',1)="","",INDEX(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),'+(i+1)+',1)),"")')
      .setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle')
      .setFontColor(CFG.COLOR.TEXT_DARK).setFontWeight('bold');
    sh.getRange(r, 10).setFormula(
      '=IFERROR(INDEX(QUERY('+TK+'!A2:Q'+MR+',"select D, sum(Q) where C = \'pengeluaran\' group by D order by sum(Q) desc limit 5 label sum(Q) \'\'",0),'+(i+1)+',2),"")')
      .setBackground('#FFFFFF').setNumberFormat('"Rp"#,##0').setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontSize(10)
      .setFontColor(CFG.COLOR.NEG);
    sh.getRange(r, 11, 1, 3).merge().setFormula(
      '=IFERROR(IF(OR(J'+r+'="",J'+r+'<=0,MAX($J$49:$J$53)<=0),"",REPT("▮",ROUND(J'+r+'/MAX($J$49:$J$53)*16,0))),"")')
      .setBackground('#FFFFFF').setFontColor(CFG.COLOR.NEG).setFontSize(11)
      .setVerticalAlignment('middle');
    sh.setRowHeight(r, 22);
    if (i % 2 === 1) sh.getRange(r, 8, 1, 6).setBackground('#FAFBFC');
  }
  sh.getRange('H48:M53').setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);

  // ===========================================================
  //  FOOTER
  // ===========================================================
  sh.setRowHeight(57, 16);
  sh.getRange('B58:M58').merge().setFormula(
    '="Catatan Audit: "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*CRITICAL*")&" kritis, "&COUNTIF('+CFG.SHEET.AUDIT+'!D2:D'+MR+',"*WARNING*")&" peringatan   •   ⚙️ Klik menu Keuangan → 🔄 Refresh setelah input transaksi   •   Filter di atas mengubah seluruh angka dashboard"')
    .setFontColor(CFG.COLOR.MUTED).setFontSize(9)
    .setBackground('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,CFG.COLOR.BORDER,SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(58, 30);
}

// ---------------------------------------------------------------------------
//  Helper: Premium KPI card  (accent stripe + label + value besar + subtitle + sparkline)
// ---------------------------------------------------------------------------
function premiumKPI_(sh, r, c, label, subtitle, formulaValue, accent, sparklineRange) {
  // Row r   : accent stripe (h:4)
  // Row r+1 : label (h:18)
  // Row r+2 : value (h:38, BIG)
  // Row r+3 : subtitle (h:18)
  // Row r+4 : sparkline (h:30, optional empty)

  // Accent stripe
  sh.getRange(r, c, 1, 3).merge().setBackground(accent);
  sh.setRowHeight(r, 4);

  // Label
  sh.getRange(r+1, c, 1, 3).merge().setValue(label).setFontSize(9).setFontColor(CFG.COLOR.MUTED)
    .setFontWeight('bold').setBackground('#FFFFFF').setVerticalAlignment('bottom')
    .setHorizontalAlignment('left').setFontFamily('Calibri');

  // Value (BIG)
  sh.getRange(r+2, c, 1, 3).merge().setFormula(formulaValue).setFontSize(20).setFontWeight('bold')
    .setFontColor(CFG.COLOR.TEXT_DARK).setBackground('#FFFFFF')
    .setNumberFormat('"Rp"#,##0;[Red]-"Rp"#,##0')
    .setVerticalAlignment('middle').setHorizontalAlignment('left').setFontFamily('Calibri');

  // Subtitle
  sh.getRange(r+3, c, 1, 3).merge().setValue(subtitle).setFontSize(9).setFontColor(CFG.COLOR.SUBMUTED)
    .setBackground('#FFFFFF').setVerticalAlignment('top')
    .setHorizontalAlignment('left').setFontFamily('Calibri');

  // Sparkline atau spacer
  if (sparklineRange) {
    sh.getRange(r+4, c, 1, 3).merge().setFormula(
      '=IFERROR(SPARKLINE(' + sparklineRange + ', {"charttype","column";"color","' + accent + '";"empty","zero";"nan","convert"}),"")')
      .setBackground('#FFFFFF').setVerticalAlignment('middle').setHorizontalAlignment('center');
  } else {
    sh.getRange(r+4, c, 1, 3).merge().setValue('').setBackground('#FFFFFF');
  }

  // Border halus (bottom + sides), tidak top karena ada stripe
  sh.getRange(r+1, c, 4, 3).setBorder(false, true, true, true, false, false, CFG.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sh.setRowHeight(r+1, 18); sh.setRowHeight(r+2, 38); sh.setRowHeight(r+3, 18); sh.setRowHeight(r+4, 30);
}

// ---------------------------------------------------------------------------
//  Helper: Section header  (judul + subtitle kecil + accent bar kiri)
// ---------------------------------------------------------------------------
function sectionHeader_(sh, rangeA1, title, subtitle, accent) {
  var rng = sh.getRange(rangeA1);
  rng.merge().setBackground('#FFFFFF')
    .setFormula('="    "&"' + title + '"&"      "&CHAR(10)&"    "&"' + subtitle + '"')
    .setFontFamily('Calibri').setVerticalAlignment('middle').setHorizontalAlignment('left');

  // Pakai rich text untuk style berbeda
  var richText = SpreadsheetApp.newRichTextValue()
    .setText('  ' + title + '   ' + subtitle)
    .setTextStyle(0, 2 + title.length + 1,
      SpreadsheetApp.newTextStyle().setBold(true).setFontSize(11).setForegroundColor(CFG.COLOR.TEXT_DARK).build())
    .setTextStyle(2 + title.length + 1, 2 + title.length + 3 + subtitle.length,
      SpreadsheetApp.newTextStyle().setBold(false).setFontSize(9).setForegroundColor(CFG.COLOR.SUBMUTED).build())
    .build();
  rng.setRichTextValue(richText);
  rng.setBorder(false, true, true, false, false, false, CFG.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  // accent border-left
  rng.setBorder(null, null, null, null, null, null);
  sh.getRange(rangeA1.split(':')[0]).setBorder(false, false, false, false, false, false);
  // pasang border kiri tebal accent dengan trick: gunakan range awal saja
  var firstCell = sh.getRange(rangeA1.split(':')[0]);
  firstCell.setBorder(null, true, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

// ---------------------------------------------------------------------------
//  Helper: Table header cell
// ---------------------------------------------------------------------------
function tableHeader_(sh, r, c, span, label, align) {
  var rng = (span > 1) ? sh.getRange(r, c, 1, span).merge() : sh.getRange(r, c);
  rng.setValue(label).setBackground('#F8FAFC').setFontWeight('bold').setFontSize(9)
    .setFontColor(CFG.COLOR.MUTED).setVerticalAlignment('middle')
    .setHorizontalAlignment(align || 'left')
    .setBorder(true, true, true, true, false, false, CFG.COLOR.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function filterDV_(ss, sh, cell, srcSheet, srcCol) {
  var src = ss.getSheetByName(srcSheet).getRange(srcCol + "2:" + srcCol + "60");
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
