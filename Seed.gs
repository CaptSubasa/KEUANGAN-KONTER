/**
 * ============================================================================
 *  Seed.gs  -  Isi data master awal + setup sheet _kalkulasi
 * ============================================================================
 */

function seedMasterData(ss) {
  // ---------- data_master (banyak blok, tiap kolom 1 daftar) ----------
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

  // ---------- akun / dompet ----------
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

  // ---------- produk (master + stok awal) ----------
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);
  // [Kode, Nama, Kategori, Satuan, HargaModalDefault, HargaJualDefault, StokAwal]
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
  // Minimal Stok (kolom M / 13)
  var minStok = produk.map(function(){ return [10]; });
  pr.getRange(2, 13, minStok.length, 1).setValues(minStok);
}

/** Tulis satu daftar ke kolom tertentu mulai baris 2 */
function setCol_(sh, col, arr) {
  var vals = arr.map(function(v){ return [v]; });
  sh.getRange(2, col, vals.length, 1).setValues(vals);
}

// ---------------------------------------------------------------------------
//  _kalkulasi : status final terpusat + named range + area helper periode
// ---------------------------------------------------------------------------
function setupKalkulasi(ss) {
  var k = ss.getSheetByName(CFG.SHEET.KALKULASI);
  k.clear();

  // A: daftar status final (FILTER dari data_master) -> dipakai semua formula
  k.getRange('A1').setValue('Status Final').setFontWeight('bold');
  k.getRange('A2').setFormula(
    '=IFERROR(FILTER(' + CFG.SHEET.DATA_MASTER + '!A2:A' + CFG.MAX_ROW + ', ' +
    CFG.SHEET.DATA_MASTER + '!B2:B' + CFG.MAX_ROW + '="Ya"),"")'
  );

  // Named range FINAL_STATUS (dipakai di banyak formula tanpa hardcode)
  var rng = k.getRange('A2:A60');
  ss.setNamedRange('FINAL_STATUS', rng);

  // C: catatan area (helper periode dashboard ditulis di Dashboard.gs)
  k.getRange('C1').setValue('(area helper dashboard - lihat Dashboard.gs)')
    .setFontColor(CFG.COLOR.MUTED);
}
