# README

## Deskripsi Proyek

Selamat datang di proyek **Konverter Excel ke KML**! Aplikasi ini dirancang untuk memudahkan pengguna dalam mengonversi data spreadsheet yang berisi informasi titik BTS (Base Transceiver Station) ke dalam format KML (Keyhole Markup Language). KML adalah format file yang digunakan untuk menampilkan data geospasial di berbagai aplikasi peta.

### Fitur Utama

- **Pengunggahan Data**: Memungkinkan pengguna untuk mengunggah file Excel yang berisi koordinat BTS, nama situs, dan informasi terkait lainnya.
- **Visualisasi Data**: Menghasilkan file KML yang dapat dibuka di aplikasi peta seperti Google Earth, menampilkan setiap BTS sebagai marker dengan deskripsi terperinci.
- **Input Dinamis**: Pengguna dapat menambahkan data baru langsung dari antarmuka tanpa harus mengedit file Excel secara manual.
- **Penyimpanan Data**: Tersedia opsi untuk menyimpan data yang telah dimodifikasi kembali ke file Excel baru.
  
### Cara Menggunakan

1. **Instalasi**: Pastikan Anda telah menginstal pustaka yang diperlukan:
   ```
   pip install pandas simplekml PyQt5
   ```

2. **Menjalankan Aplikasi**: Jalankan aplikasi dengan perintah:
   ```bash
   python main.py
   ```

3. **Muat File Excel**: Klik pada tombol "Pilih File Excel" untuk mengunggah file yang berisi data BTS.

4. **Tambahkan Data**: Gunakan antarmuka untuk menambahkan data baru jika diperlukan.

5. **Hasilkan KML**: Setelah semua data dimasukkan, klik tombol "Hasilkan File KML" untuk membuat file KML yang siap untuk digunakan.

6. **Simpan Data**: Jangan lupa untuk menyimpan data yang telah diperbarui ke dalam format Excel jika diperlukan.

### Contoh Keluaran

Setelah mengonversi data, file KML Anda akan berisi marker yang mewakili titik BTS beserta informasi relevan lainnya. Anda dapat membuka file KML ini menggunakan aplikasi mendukung KML seperti Google Earth.

### Kontribusi

Jika Anda ingin menyumbangkan fitur atau memperbaiki bug, silakan ajukan pull request atau buka issue di repositori ini.

### Lisensi

Proyek ini dilisensikan di bawah [MIT License](LICENSE).

---

Terima kasih telah menggunakan **Konverter Excel ke KML**! Nikmati eksplorasi data geospasial yang lebih baik!
