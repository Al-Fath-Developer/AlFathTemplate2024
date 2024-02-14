### Manual Penggunaan Aplikasi Pengambilan Data Google Drive

Selamat datang di Manual Penggunaan Aplikasi Pengambilan Data Google Drive. Panduan ini akan membantu Anda dalam menginstal dan menggunakan aplikasi untuk mengambil data dari Google Drive dan menyimpannya ke dalam Google Sheets.

---

### Langkah 1: Buka Spreadsheet Google Anda

1. Buka Google Sheets di peramban web Anda.

### Langkah 2: Buka Script Editor

1. Pilih menu "Extensions" -> "Apps Script".
2. Akan terbuka jendela baru dengan editor skrip.

### Langkah 3: Salin dan Tempel Kode

1. Salin kode yang ada di file ScrappingWithPrompt.js

2. Tempelkan kode yang telah disalin ke dalam editor skrip.

### Langkah 4: Simpan

1. Tekan Ctrl S untuk menyimpan kode tersebut

### Langkah 5: Berikan Izin Akses

1. Kembali ke spreadsheet Google Anda. dan coba reload spreadsheet tersebut
2. Klik menu "Scrapping" jika sudah muncul (ada di atas).
3. Pilih "Mode Simple" atau "Mode Per Folder".
4. Jika Google meminta izin akses, klik tombol "Review Permissions" atau "Authorize" (atau serupa).
5. Pilih akun Google yang ingin Anda gunakan untuk mengakses Google Drive.
6. Kalau dikatakan tidak aman, cari tulisan kecil di kiri bawah yang berwarna abu abu. lalu tekan yang ada tulisan continue atau lanjutkan
6. Klik "Allow" atau "Izinkan" untuk memberikan izin akses yang diperlukan.
7. Kembali ke spreadsheet setelah memberikan izin.

### Langkah 6: Jalankan Aplikasi

#### Mode Simple:

1. Tutup jendela editor skrip.
2. Kembali ke spreadsheet Google Anda.
3. Klik menu "Scrapping" -> "Mode Simple".
5. Pastikan bahwa di akhir baris spreadsheet Anda terdapat tulisan "Done"., karena bisa jadi ada problem karena jumlah eksekusi programnya sudah mencapai batas maksimum, sehingga di baris paling bawah tidak ada tulisan done (belum selesai proses scrapping nya) Solusinya bisa dengan melakukan run ulang mode simple


#### Mode Per Folder:

1. Tutup jendela editor skrip.
2. Kembali ke spreadsheet Google Anda.
3. Klik menu "Scrapping" -> "Mode Per Folder".

### Catatan Penting:

- **Mode Simple**: Aplikasi akan mengambil data dari semua folder di Google Drive Anda. Pastikan untuk menjalankan ulang jika belum ada tulisan "Done" di akhir baris spreadsheet.
- **Mode Per Folder**: Anda dapat memilih folder tertentu yang ingin Anda ambil datanya.

---

Anda telah berhasil menginstal dan menggunakan aplikasi untuk mengambil data dari Google Drive ke dalam Google Sheets. Jika Anda mengalami masalah atau memiliki pertanyaan, jangan ragu untuk bertanya yaaa. Terima kasih atas perhatiannya!

Thanks For ChatGPT telah membuantu membuat manual ini