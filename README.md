# Orbit | Catatan Hutang

Aplikasi web pribadi untuk mencatat hutang dari HP atau laptop. Google Apps Script menjalankan aplikasinya, sementara Google Sheets menyimpan data. Laptop tidak perlu menyala; aplikasi membutuhkan koneksi internet dan akses Google.

## Fitur

- Ringkasan total hutang, pembayaran, dan sisa tagihan.
- Riwayat di website dikelompokkan per nama, dengan total/sisa gabungan dan rincian setiap transaksi tetap dapat dibuka.
- Lunasi semua transaksi milik satu nama sekaligus; pembayaran tetap dicatat ke setiap transaksi dan tanggal pelunasan tersimpan.
- Catat transaksi dengan tanggal, nama, barang, jumlah, harga satuan, pembayaran, dan catatan.
- Cari riwayat berdasarkan nama, barang, atau catatan; edit dan hapus transaksi.
- Simpan daftar pelanggan dan barang untuk pencarian/autocomplete, dengan harga beli/modal dan harga jual. Harga terisi otomatis dan tetap bisa disesuaikan per transaksi.
- Catat belanja barang aktual beserta tanggal, jumlah, harga beli, pemasok, dan catatan; data bisa dicari, diedit, dan dihapus.
- Ubah data pelanggan/barang dari halaman Data pilihan.
- Tombol **Buka sheet** dari website membuka spreadsheet Google; tab `Data Lengkap` meratakan data transaksi dan barang agar mudah diperiksa atau diunduh untuk backup.
- Spreadsheet Google punya tab Ringkasan, Hutang, Detail Hutang, Pembayaran, Belanja, Pelanggan, Barang, Data Lengkap, serta laporan terpisah per bulan (`YYYY-MM`).
- Layout responsif untuk layar HP.

## Menyiapkan aplikasi

1. Buka [script.google.com](https://script.google.com) dengan akun Google yang akan dipakai, lalu buat project baru.
2. Aktifkan **Project Settings > Show "appsscript.json" manifest file in editor**.
3. Salin isi `apps-script/Code.gs` ke file `Code.gs`, isi `apps-script/Index.html` ke file HTML bernama `Index`, dan isi `apps-script/appsscript.json` ke manifest project.
4. Pilih fungsi `setupSpreadsheet` di editor Apps Script lalu tekan **Run**. Tinjau dan izinkan akses Google Sheets. Fungsi ini membuat spreadsheet di Google Drive dan menyimpan ID-nya untuk aplikasi.
5. Buka **Deploy > New deployment > Web app**. Pilih **Execute as: Me** dan **Who has access: Only myself**. Deploy, lalu buka URL web app di HP menggunakan akun Google yang sama.

Jangan pilih akses publik. Pembatasan "Only myself" adalah gerbang login Google; tidak ada form login atau password buatan sendiri. Jika deployment menampilkan opsi manifest sebagai `MYSELF`, tetap pastikan pengaturan **Who has access** di deployment hanya untuk akunmu.

## Spreadsheet

`setupSpreadsheet()` membuat file **Orbit | Catatan Hutang** di Google Drive akunmu. Jika ID spreadsheet sudah tersimpan, fungsi yang sama memperbarui struktur tanpa menghapus data lama. Jalankan lagi fungsi ini setelah memperbarui kode Apps Script agar tab dan kolom baru disiapkan.

- `Ringkasan`: total hutang, pembayaran, dan saldo saat ini.
- `Hutang` dan `Detail Hutang`: satu catatan per transaksi beserta rincian barang, harga jual, harga beli/modal, dan laba kotor.
- `Pembayaran`: riwayat perubahan pembayaran berdasarkan tanggal pembayaran.
- `Belanja`: pembelian barang yang benar-benar dicatat, terpisah dari modal barang yang sudah terjual.
- `Pelanggan` dan `Barang`: daftar autocomplete yang bisa dicari dan diubah dari website.
- `Data Lengkap`: satu baris per barang dalam transaksi, dengan nama/kontak pelanggan, tanggal, status, total pembayaran/sisa, harga jual/beli, modal, dan laba. Buka dari tombol **Buka sheet** di header website.
- Tab `YYYY-MM`: penjualan dan laba kotor bulan itu, modal barang terjual, pembayaran yang diterima bulan itu, belanja aktual, sisa hutang akhir bulan, serta daftar hutang dari bulan sebelumnya yang masih tersisa.

**Cara membaca angka:** laba kotor = subtotal harga jual dikurangi total harga beli/modal barang yang terjual. Belanja barang adalah kas untuk pembelian stok yang kamu catat di form Belanja. Keduanya berbeda; laporan belum mengurangi biaya lain seperti ongkir atau biaya operasional, jadi laba kotor bukan laba bersih. Pembayaran awal saat transaksi memakai tanggal transaksi; pembayaran/cicilan berikutnya dicatat pada tanggal saat perubahan disimpan.

**Data lama:** saat migrasi, harga jual katalog lama dipertahankan dan harga beli katalog lama diisi `Rp0` karena data modal sebelumnya belum ada. Lengkapi harga beli di Data pilihan, lalu edit transaksi lama dan isi harga belinya agar laba historis akurat. Pembayaran lama dipindahkan ke `Pembayaran` dengan tanggal transaksi sebagai perkiraan karena tanggal pembayaran aslinya tidak tersimpan.

## Memperbarui Apps Script yang sudah dipakai

1. Salin kode terbaru `apps-script/Code.gs` dan `apps-script/Index.html` ke project Apps Script yang sudah dipakai, menggantikan isi file lama.
2. Simpan file, pilih fungsi `setupSpreadsheet`, lalu tekan **Run** dan izinkan akses jika diminta. Spreadsheet yang sudah ada akan dimigrasikan; data hutang tidak dibuat ulang.
3. Buka **Deploy > Manage deployments**, klik ikon pensil pada web app, pilih **New version**, lalu tekan **Deploy**. URL lama tetap bisa digunakan.
4. Muat ulang web app. Pastikan akses deployment tetap **Only myself**.

Jika ingin membuka project ini dengan Apps Script lokal, buat project dari `script.google.com` lalu salin tiga file dalam folder `apps-script` ke file project yang sesuai. Deployment pertama akan meminta izin akun Google. Simpan link spreadsheet dan lakukan backup dari Google Drive secara berkala.
