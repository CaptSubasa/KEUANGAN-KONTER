/**
 * ============================================================================
 *  Ledger.gs  -  rebuildLedger() : bangun buku_besar dari semua sumber
 * ----------------------------------------------------------------------------
 *  Aturan:
 *   - Hanya status FINAL yang masuk ledger.
 *   - Mutasi / tarik tunai / transfer = perpindahan saldo (omzet 0).
 *   - Hanya admin fee yang menambah saldo bersih (laba).
 *   - Saldo Berjalan per akun dihitung urut tanggal, mulai dari Saldo Awal akun.
 * ============================================================================
 */

function rebuildLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Menghitung ulang buku besar...', 'Ledger', 4);

  // 1. Kumpulan status final
  var finalSet = getFinalStatusSet_(ss);

  // 2. Saldo awal tiap akun
  var balances = {};
  var akSh = ss.getSheetByName(CFG.SHEET.AKUN);
  var akRows = readRows_(akSh, 3);
  akRows.forEach(function(r){ if (r[0] !== '') balances[r[0]] = num_(r[2]); });

  // 3. Kumpulkan entri ledger dari semua sumber
  var E = [];   // {d, ref, src, jenis, akun, masuk, keluar, kat, status, ket, user}
  collectPenjualan_(ss, finalSet, E);
  collectPembelian_(ss, finalSet, E);
  collectTransaksiKas_(ss, finalSet, E);
  collectUtangPiutang_(ss, E);

  // 4. Urutkan berdasarkan tanggal lalu hitung saldo berjalan per akun
  E.sort(function(a, b){ return (a.d ? a.d.getTime() : 0) - (b.d ? b.d.getTime() : 0); });
  var out = [];
  E.forEach(function(e){
    if (!(e.akun in balances)) balances[e.akun] = 0;
    balances[e.akun] += (e.masuk - e.keluar);
    out.push([e.d, e.ref, e.src, e.jenis, e.akun, e.masuk, e.keluar, balances[e.akun],
              e.kat, e.status, e.ket, e.user]);
  });

  // 5. Tulis ke buku_besar
  var bb = ss.getSheetByName(CFG.SHEET.BUKU_BESAR);
  bb.getRange(2, 1, CFG.MAX_ROW - 1, 12).clearContent();
  if (out.length > 0) bb.getRange(2, 1, out.length, 12).setValues(out);

  SpreadsheetApp.flush();
  ss.toast('Buku besar diperbarui: ' + out.length + ' baris.', 'Ledger', 5);
}

// ---------------------------------------------------------------------------
//  Kolektor per sumber
// ---------------------------------------------------------------------------
function collectPenjualan_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.PENJUALAN);
  var rows = readRows_(sh, 21);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], kat=r[5], totHPP=num_(r[9]), totJual=num_(r[10]),
        metode=r[12], akunMasuk=r[13], akunModal=r[14], status=r[15], user=r[18];
    if (tgl==='' || !finalSet[status]) return;
    // Uang masuk (kecuali penjualan kredit -> jadi piutang, dicatat terpisah)
    if (metode !== 'Kredit' && akunMasuk !== '') {
      E.push(entry_(tgl, id, CFG.SHEET.PENJUALAN, 'Penjualan', akunMasuk, totJual, 0, kat, status, 'Omzet penjualan', user));
    }
    // HPP keluar dari Akun Modal / Sumber HPP (saldo supplier berkurang)
    if (akunModal !== '' && totHPP > 0) {
      E.push(entry_(tgl, id, CFG.SHEET.PENJUALAN, 'HPP Penjualan', akunModal, 0, totHPP, kat, status, 'Pemakaian modal/HPP', user));
    }
  });
}

function collectPembelian_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.PEMBELIAN);
  var rows = readRows_(sh, 17);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], kat=r[5], total=num_(r[8]), akunBayar=r[10],
        status=r[11], jenisBeli=r[12], user=r[15];
    if (tgl==='' || !finalSet[status]) return;
    if (jenisBeli !== 'Kredit' && akunBayar !== '') {
      E.push(entry_(tgl, id, CFG.SHEET.PEMBELIAN, 'Pembelian Stok', akunBayar, 0, total, kat, status, 'Beli stok', user));
    }
  });
}

function collectTransaksiKas_(ss, finalSet, E) {
  var sh = ss.getSheetByName(CFG.SHEET.TRANSAKSI_KAS);
  var rows = readRows_(sh, 16);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], jenis=r[2], kat=r[3], akunK=r[4], akunM=r[5],
        nominal=num_(r[6]), fee=num_(r[7]), status=r[10], arah=r[11], pihak=r[12], user=r[14];
    if (tgl==='' || !finalSet[status]) return;

    switch (jenis) {
      case 'pengeluaran':
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Pengeluaran',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'modal':
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Modal Masuk',akunM,nominal,0,kat,status,pihak,user));
        break;
      case 'mutasi':
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Mutasi (keluar)',akunK,0,nominal,kat,status,pihak,user));
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Mutasi (masuk)',akunM,nominal,0,kat,status,pihak,user));
        if (fee > 0 && akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Biaya Admin Mutasi',akunK,0,fee,kat,status,pihak,user));
        break;
      case 'tarik_tunai':
        // pelanggan transfer ke rekening kita (pokok+fee), kita beri tunai (pokok)
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Tarik Tunai (masuk)',akunM,nominal+fee,0,kat,status,pihak,user));
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Tarik Tunai (tunai keluar)',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'transfer_uang':
        // pelanggan beri tunai (pokok+fee), kita transfer dari rekening (pokok)
        if (akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Transfer Uang (tunai masuk)',akunM,nominal+fee,0,kat,status,pihak,user));
        if (akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Transfer Uang (keluar)',akunK,0,nominal,kat,status,pihak,user));
        break;
      case 'koreksi_kas':
        if (arah === 'Tambah' && akunM !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Koreksi Kas (+)',akunM,nominal,0,kat,status,pihak,user));
        else if (arah === 'Kurang' && akunK !== '') E.push(entry_(tgl,id,CFG.SHEET.TRANSAKSI_KAS,'Koreksi Kas (-)',akunK,0,nominal,kat,status,pihak,user));
        break;
    }
  });
}

function collectUtangPiutang_(ss, E) {
  var sh = ss.getSheetByName(CFG.SHEET.UTANG_PIUTANG);
  var rows = readRows_(sh, 13);
  rows.forEach(function(r){
    var tgl=r[0], id=r[1], arah=r[2], pihak=r[3], dibayar=num_(r[6]), akunBayar=r[8];
    if (tgl==='' || dibayar<=0 || akunBayar==='') return;
    if (arah === 'Piutang') {
      E.push(entry_(tgl,id,CFG.SHEET.UTANG_PIUTANG,'Pelunasan Piutang',akunBayar,dibayar,0,'Piutang','-',pihak,''));
    } else if (arah === 'Hutang') {
      E.push(entry_(tgl,id,CFG.SHEET.UTANG_PIUTANG,'Pembayaran Hutang',akunBayar,0,dibayar,'Hutang','-',pihak,''));
    }
  });
}

// ---------------------------------------------------------------------------
//  Util
// ---------------------------------------------------------------------------
function entry_(d, ref, src, jenis, akun, masuk, keluar, kat, status, ket, user) {
  return { d:(d instanceof Date ? d : new Date(d)), ref:ref, src:src, jenis:jenis,
           akun:akun, masuk:masuk, keluar:keluar, kat:kat, status:status, ket:ket, user:user };
}

function getFinalStatusSet_(ss) {
  var dm = ss.getSheetByName(CFG.SHEET.DATA_MASTER);
  var rows = readRows_(dm, 2);
  var set = {};
  rows.forEach(function(r){ if (r[0] !== '' && String(r[1]).toLowerCase() === 'ya') set[r[0]] = true; });
  return set;
}

function readRows_(sh, nCol) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, nCol).getValues();
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}
