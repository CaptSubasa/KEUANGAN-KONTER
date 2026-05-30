# Sistem Manajemen Keuangan Konter Pulsa / Digital Payment

Generator **Google Apps Script** yang membangun otomatis satu workbook Google Sheets lengkap (16 sheet) untuk mengelola keuangan usaha konter pulsa: penjualan, stok, kas/bank, piutang, hutang, modal, mutasi, tarik tunai, transfer, koreksi, buku besar, laporan laba rugi, dashboard, audit, dan pengujian.

Dirancang **siap dikembangkan ke web app** — nama sheet = `snake_case` (calon nama tabel/endpoint), ada ID unik, dan logika status terpusat.

---

## 1. Cara Pasang (sekali saja)

1. Buat **Google Sheets** baru (kosong).
2. Menu **Extensions / Ekstensi → Apps Script**.
3. Hapus isi `Code.gs` bawaan. Lalu **buat file** untuk tiap file `.gs` di folder ini dan salin isinya:
   - `Code.gs`, `Structure.gs`, `Seed.gs`, `Formulas.gs`, `Ledger.gs`, `Validation.gs`, `Dashboard.gs`, `Protection.gs`, `Triggers.gs`, `DummyData.gs`
   - (opsional) ganti manifest dengan `appsscript.json`.
4. Pilih fungsi **`buildAll`** di toolbar Apps Script, klik **Run**.
5. Saat diminta, **izinkan otorisasi** (wajib agar script bisa menulis sheet).
6. Tunggu ±1–2 menit. Setelah selesai, kembali ke Google Sheets dan **reload**.
7. Muncul menu **⚙️ Keuangan** di bar menu.

---

## 2. Struktur 16 Sheet

| Kelompok | Sheet | Fungsi |
|---|---|---|
| Inti | `penjualan` | Penjualan pulsa/data/voucher/token/e-wallet/layanan |
| Inti | `pembelian` | Pembelian stok dari supplier |
| Inti | `produk` | Master produk + perhitungan stok |
| Inti | `transaksi_kas` | Pengeluaran, modal, mutasi, tarik tunai, transfer, koreksi kas (kolom `Jenis`) |
| Inti | `koreksi_stok` | Penyesuaian stok (berlaku jika Approved) |
| Inti | `utang_piutang` | Piutang + hutang (kolom `Arah`) |
| Inti | `buku_besar` | Buku besar kas/bank — **sumber kebenaran saldo** |
| Referensi | `akun` | Daftar akun/dompet + saldo |
| Referensi | `data_master` | Status, kategori, jenis, supplier, pelanggan, user |
| Sistem (hidden) | `_dropdown` | Sumber semua dropdown |
| Sistem (hidden) | `_kalkulasi` | Mesin hitung + data grafik |
| Tampilan | `dashboard` | KPI, grafik dinamis, alert, ringkasan |
| Laporan | `laporan` | Laba rugi bulanan (pilih tahun) |
| Audit | `audit` | 15 pemeriksaan otomatis |
| Audit | `skenario_uji` | 20 skenario pengujian |
| Audit | `log_perubahan` | Catatan perubahan |

---

## 3. Aturan Logika Penting

- **Status final** = `Lunas, Selesai, Sukses, Approved, Disetujui`. Hanya status final yang memengaruhi stok, saldo, laba, dan buku besar. Status `Pending/Draft/Proses/Batal/Gagal` diabaikan.
- Daftar status final diatur terpusat di `data_master` kolom **Final? (Ya/Tidak)** → ubah di sana, seluruh sistem ikut.
- **Omzet hanya dari `penjualan`.** Mutasi, tarik tunai, dan transfer **bukan omzet** — hanya **admin fee** yang jadi laba.
- **Snapshot harga**: saat memilih Nama Produk di `penjualan`, Harga Modal & Jual otomatis dibekukan jadi nilai statis (lewat `onEdit`). Perubahan harga master TIDAK mengubah transaksi lama.
- `Laba Bersih = Laba Kotor + Admin Fee − Pengeluaran`.

---

## 4. Cara Pakai Harian

1. **Tambah produk** baru di sheet `produk` (Kode, Nama, Kategori, harga, stok awal, minimal stok).
2. **Input penjualan** di `penjualan`: isi Tanggal, pilih Nama Produk (harga terisi otomatis), Qty, Metode, Akun Uang Masuk, Status, dll.
3. **Input pembelian/pengeluaran/dll** di sheet terkait.
4. Klik menu **⚙️ Keuangan → 🔄 Refresh Buku Besar & Saldo** setiap selesai input agar saldo & dashboard ter-update.
5. Pantau **dashboard** (filter Tahun/Bulan/Mode Waktu/Akun/Kategori) dan **audit**.

> Catatan: buku besar dihitung oleh script (bukan formula hidup), jadi **tekan Refresh** setelah input agar saldo akun & KPI ikut terbarui.

---

## 5. Menu ⚙️ Keuangan

- **🔄 Refresh Buku Besar & Saldo** — bangun ulang `buku_besar` dari semua sumber + hitung saldo berjalan.
- **🧪 Jalankan Skenario Uji** — refresh + ringkas hasil PASS/FAIL.
- **🏗️ BUILD ULANG SEMUA (reset)** — bangun ulang seluruh sistem (menghapus isi).
- **🔓 Lepas Semua Proteksi** — hapus proteksi warning bila perlu edit bebas.

---

## 6. Catatan Pengembangan ke Web App

- Setiap sheet inti = kandidat tabel/endpoint (`/api/penjualan`, `/api/transaksi_kas`, ...).
- `transaksi_kas` bersifat polymorphic (kolom `Jenis`) → satu endpoint, validasi kondisional per jenis.
- `rebuildLedger()` bisa dipromosikan jadi job backend; fungsi `doGet/doPost` Apps Script dapat ditambahkan untuk menjadikannya Web App/REST API tanpa migrasi data.

---

## 7. Status / Hal yang Masih Bisa Disempurnakan

- Grafik mode **Mingguan** memakai pembagian 7 harian sederhana (belum ISO week).
- Proteksi memakai mode *warning-only* (tidak memblokir) agar input harian tidak terganggu.
- Beberapa skenario dashboard ditandai **MANUAL** (perlu cek mata).

Silakan jalankan `buildAll`, lalu kabari bagian mana yang ingin diperbaiki/ditambah.
