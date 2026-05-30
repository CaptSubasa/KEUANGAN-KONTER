/**
 * ============================================================================
 *  Triggers.gs  -  onEdit: bekukan Harga Modal & Jual Snapshot (Opsi C)
 * ----------------------------------------------------------------------------
 *  Saat Nama Produk dipilih di sheet penjualan, harga modal & jual produk
 *  ditulis sebagai NILAI STATIS (bukan formula) ke kolom H & I baris itu.
 *  Sehingga perubahan harga di master TIDAK mengubah transaksi lama.
 * ============================================================================
 */

function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET.PENJUALAN) return;

  var startCol = e.range.getColumn();
  var endCol = startCol + e.range.getNumColumns() - 1;
  // Kolom D (Nama Produk) = 4
  if (4 < startCol || 4 > endCol) return;

  var startRow = e.range.getRow();
  var nRows = e.range.getNumRows();
  if (startRow < 2) startRow = 2;

  var ss = e.source;
  var prMap = buildProdukPriceMap_(ss);

  for (var i = 0; i < nRows; i++) {
    var row = e.range.getRow() + i;
    if (row < 2) continue;
    freezeSnapshotRow_(sh, prMap, row);
  }
}

/** Bekukan harga snapshot untuk satu baris penjualan */
function freezeSnapshotRow_(sh, prMap, row) {
  var nama = String(sh.getRange(row, 4).getValue()).trim();   // D
  if (nama === '') {
    sh.getRange(row, 8, 1, 2).clearContent();                 // H,I
    return;
  }
  var p = prMap[nama];
  if (!p) return;                                             // produk tak dikenal -> biarkan
  sh.getRange(row, 8).setValue(p.modal);                      // H Harga Modal Snapshot
  sh.getRange(row, 9).setValue(p.jual);                       // I Harga Jual Snapshot
}

/** Peta { namaProduk: {modal, jual} } dari sheet produk */
function buildProdukPriceMap_(ss) {
  var pr = ss.getSheetByName(CFG.SHEET.PRODUK);
  var last = pr.getLastRow();
  var map = {};
  if (last < 2) return map;
  var vals = pr.getRange(2, 2, last - 1, 5).getValues();      // B..F : nama,kategori,satuan,modal,jual
  vals.forEach(function(r){
    var nama = String(r[0]).trim();
    if (nama !== '') map[nama] = { modal: num_(r[3]), jual: num_(r[4]) };
  });
  return map;
}

/** Bekukan ulang semua baris penjualan yang snapshot-nya kosong (manual/menu) */
function freezeAllSnapshots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.SHEET.PENJUALAN);
  var last = sh.getLastRow();
  if (last < 2) return;
  var prMap = buildProdukPriceMap_(ss);
  var data = sh.getRange(2, 4, last - 1, 6).getValues();      // D..I
  for (var i = 0; i < data.length; i++) {
    var nama = String(data[i][0]).trim();
    var hModal = data[i][4], hJual = data[i][5];               // H,I
    if (nama !== '' && (hModal === '' || hJual === '')) {
      freezeSnapshotRow_(sh, prMap, i + 2);
    }
  }
}
