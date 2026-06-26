# Rekap Excel Booking

Aplikasi lokal untuk mengubah file Excel mentah booking AYO menjadi workbook rekap dengan sheet:

- `Summary <bulan>`
- `Walk In <bulan>`
- `AYO-<bulan>`
- `ALL-<bulan>`

Dashboard juga menyediakan modul:

- `Omset Lapangan`
- `Omset Perkategori Olsera`
- `PDF TO EXCEL` untuk mengekspor tabel/teks beberapa PDF ke satu workbook, dengan sheet ringkasan dan satu sheet per halaman

PDF hasil scan yang hanya berisi gambar perlu diproses OCR terlebih dahulu.

## Jalankan Mode Desktop

```powershell
python desktop_app.py
```

Isi:

- File Excel mentah
- Mode `Harian` atau `Sebulan`
- Tanggal, contoh `8` untuk mode harian
- Bulan, contoh `Juni`
- Tahun opsional, contoh `2026`

File mentah yang diupload akan disimpan ke database lokal `data/rekap.db`. Setelah data sudah ada di database, kamu bisa proses ulang tanggal/bulan yang sama tanpa upload file lagi.

Jika besok upload data tanggal 9, database tetap menyimpan data tanggal 8 dan tanggal 9. Pilih tanggal `9` dan bulan `Juni` untuk membuat rekap tanggal 9.

Setiap proses rekap selesai akan masuk ke log database. Di aplikasi desktop, tabel `Log Rekap Selesai` menampilkan tanggal yang sudah diproses, total, dan file output yang bisa dibuka ulang.

Untuk data sebulan, pilih mode `Sebulan`, pilih bulan, lalu proses. Sheet `Summary` akan otomatis mengisi total per tanggal dalam bulan tersebut.

## Build EXE

```powershell
.\build_exe.ps1
```

Script akan membuat versi Windows 64-bit dan 32-bit sekaligus. Pada build
pertama, Python 32-bit dan dependensi build akan disiapkan secara otomatis.

Hasil EXE:

```text
dist\RekapExcelBooking-x64.exe
dist\RekapExcelBooking-x86.exe
```

- Kirim versi `x64` untuk Windows 64-bit.
- Kirim versi `x86` untuk Windows 32-bit. Versi ini juga dapat berjalan pada
  Windows 64-bit, tetapi versi `x64` tetap disarankan untuk PC 64-bit.
- Setiap ada perubahan aplikasi, cukup jalankan kembali `build_exe.ps1` agar
  kedua EXE diperbarui dari source code yang sama.

## Aturan Pemisahan

- Masuk `Walk In` jika `Booking ID` diawali `MN/`, atau `Payment Method` berisi `MANUAL`/`WALK`.
- Selain itu masuk `AYO`.
- `ALL` adalah gabungan semua data valid, disortir per court lalu jam booking.
- `Summary` menghitung `Revenue Venue` per tanggal, court, dan channel.
- Nomor kolom `No` dibuat ulang berurutan setelah sorting.
- Jika ada `Booking ID` duplikat di file input, semua baris tetap dihitung. Aplikasi hanya memberi peringatan duplikat.
- Upload ulang data pada tanggal yang sama akan mengganti snapshot tanggal tersebut di database lokal.
- Booking ID yang tercatat sebagai payment fail di database pengecualian tidak ikut rekap.
