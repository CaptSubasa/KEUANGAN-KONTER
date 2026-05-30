/**
 * ============================================================================
 *  DummyData.gs  -  Data contoh untuk testing + 20 skenario uji
 * ============================================================================
 */

function insertDummyData(ss) {
  var now = new Date();
  var Y = now.getFullYear();
  var cm = now.getMonth() + 1;
  var months = [];
  for (var mm = Math.max(1, cm - 2); mm <= cm; mm++) months.push(mm);

  var prMap = buildProdukPriceMap_(ss);

  // ---------------- penjualan ----------------
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

  // ---------------- pembelian ----------------
  var buys = [
    mkBuy_(Y,months[0],3,'Supplier Pulsa A','Pulsa 10K',50,10500,'Transfer Bank','Bank BCA','Selesai','Tunai','Topup awal','Owner'),
    mkBuy_(Y,months[months.length-1],4,'Distributor Data B','Paket Data 1GB',40,9000,'Transfer Bank','Bank BCA','Selesai','Tunai','Stok data','Owner'),
    mkBuy_(Y,cm,2,'Agen Token D','Token Listrik 20K',30,20500,'Tunai','Kas Utama','Lunas','Tunai','Stok token','Kasir 1'),
    mkBuy_(Y,cm,6,'PPOB Center C','Voucher Game 10K',20,9500,'Kredit','Saldo Supplier','Selesai','Kredit','Beli kredit','Owner'),
    mkBuy_(Y,cm,7,'Supplier Pulsa A','Pulsa 25K',20,24800,'Transfer Bank','Bank BCA','Pending','Tunai','Belum diterima','Owner') // pending: tidak nambah stok
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.PEMBELIAN), buys,
    {A:'tgl',C:'supplier',D:'produk',G:'qty',H:'harga',J:'metode',K:'akunBayar',L:'status',M:'jenisBeli',O:'catatan',P:'user'});

  // ---------------- transaksi_kas ----------------
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

  // ---------------- koreksi_stok ----------------
  var ks = [
    mkKs_(Y,cm,5,'Voucher Game 10K','Kurang',2,'Barang Rusak','Approved','Owner'),
    mkKs_(Y,cm,6,'Pulsa 10K','Tambah',3,'Stok Opname','Pending','Kasir 1') // pending: tidak diterapkan
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.KOREKSI_STOK), ks,
    {A:'tgl',C:'produk',E:'arah',F:'qty',G:'alasan',H:'status',I:'user'});

  // ---------------- utang_piutang ----------------
  var up = [
    mkUp_(Y,cm,2,'Piutang','Toko Maju',64000,dt_(Y,cm,16),0,'','Proses','Penjualan kredit Paket Data 5GB'),
    mkUp_(Y,cm,3,'Hutang','PPOB Center C',190000,dt_(Y,cm,20),0,'','Proses','Pembelian voucher kredit'),
    mkUp_(Y,months[0],10,'Piutang','Warung Bu Tini',50000,dt_(Y,months[0],15),20000,'Kas Utama','Proses','Piutang lama (telat)')
  ];
  writeDummy_(ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG), up,
    {A:'tgl',C:'arah',D:'pihak',E:'nominal',F:'jatuhTempo',G:'dibayar',I:'akunBayar',J:'status',L:'keterangan'});

  SpreadsheetApp.flush();
}

// ---------------------------------------------------------------------------
//  Pembuat baris
// ---------------------------------------------------------------------------
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

/** Tulis dummy ke kolom-kolom input (membersihkan kolom itu dulu, tidak menyentuh kolom formula) */
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

// ===========================================================================
//  SKENARIO UJI (20)
// ===========================================================================
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

// ---------------------------------------------------------------------------
//  Runner skenario (menu)
// ---------------------------------------------------------------------------
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
