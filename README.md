# Deskripsi Program
Program ini adalah aplikasi desktop yang ditulis dalam Python menggunakan Tkinter untuk antarmuka pengguna dan pdfplumber untuk ekstraksi data dari file PDF. Aplikasi ini dirancang untuk mengonversi daftar bagian yang terdapat dalam file PDF ke dalam format Excel (.xlsx). Pengguna dapat memilih file PDF yang berisi informasi tentang blok dan profil, dan program akan mengekstrak data tersebut, menyusunnya dalam format tabel, dan menyimpannya sebagai file Excel. Program ini sangat berguna bagi para perancang yang bekerja dengan data CADMATIC dan memerlukan cara yang efisien untuk mengelola informasi bagian.

# Tutorial Cara Menjalankan Program dan Deploy
## Persyaratan
Sebelum menjalankan program, pastikan telah menginstal Python dan pustaka yang diperlukan.

## Instalasi Pustaka yang Diperlukan
- Buka terminal atau command prompt.
- Jalankan perintah berikut untuk menginstal pustaka yang diperlukan:
```console
pip install pdfplumber pandas xlsxwriter
```

## Menyimpan Kode Program
- Salin kode program di atas ([PDF to Excel Converter to CADMATIC PartsList](https://github.com/NEAR07/PDF-to-Excel-Converter-to-CADMATIC-PartsList/blob/main/PDF-to-Excel-Converter-to-CADMATIC-PartsList.py)).
- Buka editor teks (seperti Notepad, VSCode, atau PyCharm) dan tempelkan kode tersebut.
- Simpan file dengan nama pdf_to_excel_converter.py.

## Menjalankan Program
- Buka terminal atau command prompt.
- Arahkan ke direktori tempat Anda menyimpan file pdf_to_excel_converter.py menggunakan perintah cd. Contoh:
```console
cd path\to\your\directory
```
- Jalankan program dengan perintah:
```console
python pdf_to_excel_converter.py
```

## Menggunakan Aplikasi
- Setelah program dijalankan, jendela aplikasi akan muncul.
- Klik tombol "Select PDF File" untuk memilih file PDF yang ingin Anda konversi.
- Pilih lokasi dan nama file untuk menyimpan file Excel yang dihasilkan.
- Program akan mengekstrak data dari PDF dan menyimpannya dalam format Excel.
- Setelah proses selesai, jendela konfirmasi akan muncul, menunjukkan lokasi file Excel yang telah disimpan.

## Men-deploy Aplikasi
Jika ingin mendistribusikan aplikasi ini kepada pengguna lain, Maka dapat menggunakan alat seperti PyInstaller untuk mengonversi skrip Python menjadi file executable (.exe).

- Instal PyInstaller:
```console
pip install pyinstaller
```
- Jalankan perintah berikut untuk membuat executable:
```console
pyinstaller --onefile pdf_to_excel_converter.py
```
Setelah proses selesai, Maka akan menemukan file executable di dalam folder dist.
