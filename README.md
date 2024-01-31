# AlFathTemplate2024
Bismilah, template ini ga secara tertutup khusus untuk kegiatan Al-Fath saja, tetapi bisa dipakai untuk hal lain yang bermanfaat

## Perizinan
### Instalasi
1. Buat Google Form untuk perizinan
2. Setelah google form jadi, pertanyaan pertama langsung tanyakan NIM dari pengisi, dan yang kedua adalah waktu perizinan (dia izin untuk kapan)m, dengan tipe data date (bukan date time, tapi date doang)
3. Untuk sisanya bisa menyesuaikan dengan kebutuhan
4. Note: pertanyaan NIM dibuat paling pertama, tetapi urutannya boleh diubah, yang penting NIM dibuat paling pertama sat pembuatan Gform dan pertanyaan waktu perizinan dibuat kedua
5. buat spreadsheet untuk hadil google form tersebut (jawaban -> buat spreadshdeet baru)
6. buka menu ekstensi, lalu pilih appscript
7. masukan kodingan dari Perizinan.js ke daalam appscript
8. ubah SPREADSHEET_ID dan SHEET_NAME dengan file spreadsheet yang anda gunakan untuk rekap data
9. simpan kodingan tersebut, lalu buka menu "trigger"  atua "pemicu" di kiri
10. tambahkan pemicu, lalu isi sebagai berikut: ![image](https://github.com/atiohaidar/AlFathTemplate2024/assets/66902140/cbd686d6-dd9c-49bb-8cef-e958773e5470)
11. simpan, lalu uji coba google form teresebut
### Konsep
Pengisia form bedasarkan NIM yang diinputkan dan waktu perizinan yang diinputkan, sisannya bisa menyesuaikan
Gunakan template Spreadsheet yang sudah ada: https://docs.google.com/spreadsheets/d/1mvaGSfS-WlA1noFGG_pe1SskxWDBnXiT3lb4pd4ub08/edit#gid=0



## Presensi
### Instalasi
1. Buat Google Form untuk presensi
2. Setelah google form jadi, pertanyaan pertama langsung tanyakan NIM dari pengisi
3. Untuk sisanya bisa menyesuaikan dengan kebutuhan
4. Note: pertanyaan NIM dibuat paling pertama, tetapi urutannya boleh diubah, yang penting NIM dibuat paling pertama sat pembuatan Gform
5. buat spreadsheet untuk hadil google form tersebut (jawaban -> buat spreadshdeet baru)
6. buka menu ekstensi, lalu pilih appscript
7. masukan kodingan dari Presensi.js ke daalam appscript
8. ubah SPREADSHEET_ID dan SHEET_NAME dengan file spreadsheet yang anda gunakan untuk rekap data
9. simpan kodingan tersebut, lalu buka menu "trigger"  atua "pemicu" di kiri
10. tambahkan pemicu, lalu isi sebagai berikut: ![image](https://github.com/atiohaidar/AlFathTemplate2024/assets/66902140/cbd686d6-dd9c-49bb-8cef-e958773e5470)
11. simpan, lalu uji coba google form teresebut
### Konsep
Rekap otomati ini bedasarkan NIM dan waktu kapan dia melakukan submit (timestamp). sisanya bisa menyesuaikan
Gunakan template Spreadsheet yang sudah ada: https://docs.google.com/spreadsheets/d/1mvaGSfS-WlA1noFGG_pe1SskxWDBnXiT3lb4pd4ub08/edit#gid=0


