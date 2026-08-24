# 🏦 Mutasi BCA — CSV & PDF ke Excel

> **Lima skrip Python untuk mengonversi, menggabungkan, dan menganalisis mutasi rekening BCA — dari file CSV atau PDF ke laporan Excel siap pakai**

Kumpulan skrip standalone yang saling melengkapi: konversi CSV mutasi BCA ke Excel (dua varian), ekstraksi PDF mutasi ke Excel, penggabungan banyak file Excel menjadi satu workbook per rekening, dan deteksi transfer antar rekening BCA secara otomatis.

---

## 📋 Daftar Isi

- [Gambaran Umum & Alur Kerja](#-gambaran-umum--alur-kerja)
- [Prasyarat](#-prasyarat)
- [Deskripsi Tiap Skrip](#-deskripsi-tiap-skrip)
  - [`bcacsv2excel.py` — Konversi CSV ke Excel](#1-bcacsv2excelpy--konversi-csv-ke-excel)
  - [`bcacsv2exceldbcr.py` — Konversi CSV + Pisah DB/CR](#2-bcacsv2exceldbcrpy--konversi-csv--pisah-dbcr)
  - [`bca_extract_pdf.py` — Ekstraksi PDF ke Excel](#3-bca_extract_pdfpy--ekstraksi-pdf-ke-excel)
  - [`gabung_BCA.py` — Gabungkan File Excel](#4-gabung_bcapy--gabungkan-file-excel)
  - [`cek_tarikan_BCA2BCA.py` — Deteksi Transfer Antar Rekening](#5-cek_tarikan_bca2bcapy--deteksi-transfer-antar-rekening)
- [Alur Kerja yang Disarankan](#-alur-kerja-yang-disarankan)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum & Alur Kerja

Setiap skrip dapat dijalankan secara **mandiri** sesuai kebutuhan, atau dikombinasikan sebagai pipeline. Tidak ada orchestrator — jalankan satu per satu.

```
Sumber Data
    │
    ├── File CSV BCA  ──→  bcacsv2excel.py        → Excel (format asli)
    │                └──→  bcacsv2exceldbcr.py    → Excel (DB & CR terpisah)
    │
    └── File PDF BCA  ──→  bca_extract_pdf.py     → Excel (Debet, Kredit, Saldo)
                                  │
                         (semua Excel dari langkah di atas)
                                  │
                          gabung_BCA.py            → Hasil_Gabungan_Mutasi_BCA.xlsx
                                  │
                  cek_tarikan_BCA2BCA.py           → Laporan_Transfer_Antar_Bank.xlsx
```

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan. Telah diuji pada **Windows 11** dan **Ubuntu 24.04**.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlsxwriter pdfplumber
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Semua skrip | Baca, transformasi, dan simpan data |
| `openpyxl` | `bcacsv2excel.py`, `bcacsv2exceldbcr.py`, `bca_extract_pdf.py` | Tulis `.xlsx` dan auto-fit kolom |
| `xlsxwriter` | `bca_extract_pdf.py`, `gabung_BCA.py`, `cek_tarikan_BCA2BCA.py` | Buat `.xlsx` dengan format angka dan lebar kolom |
| `pdfplumber` | `bca_extract_pdf.py` | Ekstraksi teks berbasis koordinat dari PDF |
| `re`, `glob`, `os`, `difflib`, `datetime` | Semua | Standard library |

### Cara menjalankan

Letakkan skrip di folder yang sama dengan file input (`*.csv` atau `*.pdf`), lalu:

```bash
# Klik dua kali (Windows), atau:
python bcacsv2excel.py
python bcacsv2exceldbcr.py
python bca_extract_pdf.py
python gabung_BCA.py
python cek_tarikan_BCA2BCA.py
```

---

## 🔍 Deskripsi Tiap Skrip

### 1. `bcacsv2excel.py` — Konversi CSV ke Excel

Mengonversi **semua file `.csv`** di folder yang sama ke format Excel (`.xlsx`) dengan nama file yang otomatis dibentuk dari isi data.

#### Cara pakai
1. Letakkan `bcacsv2excel.py` di folder yang berisi file CSV mutasi BCA.
2. Jalankan skrip.
3. File Excel akan muncul di folder yang sama.

#### Penamaan output otomatis

Nama file output dibentuk dari metadata di dalam CSV:

```
BCA {4 digit terakhir rekening} {tanggal} {bulan}.xlsx
```

Contoh: `BCA 3456 15 AGU.xlsx` atau `BCA 3456 1 - 31 AGU.xlsx` (jika periode lebih dari satu hari).

Data diambil dari:
- **Baris ke-2** CSV → 4 digit terakhir nomor rekening
- **Baris ke-4** CSV → periode tanggal (format `DD/MM/YYYY - DD/MM/YYYY`)

Jika file output sudah ada, ditambahkan suffix angka: `BCA 3456 15 AGU-1.xlsx`, `BCA 3456 15 AGU-2.xlsx`, dst.

#### Fitur
- Deteksi separator CSV otomatis (`sep=None`)
- Bersihkan tanda kutip dan koma dari tepi nilai sel
- Auto-fit lebar semua kolom
- Toleransi error per file — jika satu CSV gagal, lanjut ke berikutnya

---

### 2. `bcacsv2exceldbcr.py` — Konversi CSV + Pisah DB/CR

Varian dari `bcacsv2excel.py` yang secara otomatis **mendeteksi kolom berformat `nominal DB` atau `nominal CR`** dan memisahkannya menjadi dua kolom numerik terpisah: `DB` (debit) dan `CR` (kredit).

#### Cara pakai
1. Letakkan `bcacsv2exceldbcr.py` di folder berisi file CSV.
2. Jalankan skrip.

#### Perbedaan utama dari `bcacsv2excel.py`

| Aspek | `bcacsv2excel.py` | `bcacsv2exceldbcr.py` |
|---|---|---|
| Penamaan output | Otomatis dari metadata CSV | Sama dengan nama file CSV asli |
| Kolom DB/CR | Tetap sebagai teks asli | Dipecah menjadi dua kolom numerik |
| Input `"1.234,56 DB"` | Disimpan sebagai teks `1.234,56 DB` | Kolom `DB` = `1234.56`, Kolom `CR` = kosong |
| Input `"2.345,67 CR"` | Disimpan sebagai teks `2.345,67 CR` | Kolom `DB` = kosong, Kolom `CR` = `2345.67` |

#### Logika deteksi kolom DB/CR

Kolom dianggap ber-format DB/CR jika **setidaknya satu** dari 20 baris pertama cocok dengan pola:

```
regex: ^\s*([\d.,]+)\s*(DB|CR)\s*$
```

Kolom yang terdeteksi akan diganti dengan dua kolom baru (`DB` dan `CR`). Kolom lain tidak berubah.

---

### 3. `bca_extract_pdf.py` — Ekstraksi PDF ke Excel

Mengekstrak data transaksi dari **mutasi rekening BCA format PDF** menjadi tabel Excel terstruktur dengan kolom Tanggal, Keterangan, Debet, Kredit, dan Saldo.

#### Cara pakai
1. Letakkan `bca_extract_pdf.py` di folder berisi file PDF mutasi BCA.
2. Jalankan skrip.
3. Pilih mode:

```
1. Pindai Perdokumen  ← proses satu file yang dipilih
2. Proses Semua Dokumen  ← proses semua PDF sekaligus
Masukkan (1/2):
```

#### Penamaan output

Setiap PDF menghasilkan satu file Excel di folder yang sama:
```
{nama_file}.pdf  →  {nama_file}_Excel.xlsx
```

#### Kolom output

| Kolom | Format | Keterangan |
|---|---|---|
| `Tanggal` | `DD/MM/2025` | Tanggal transaksi (⚠️ lihat Catatan Penting) |
| `Keterangan` | Teks, wrap | Deskripsi transaksi (multi-baris digabung) |
| `Debet` | `#,##0.00` | Nilai debet (0 jika bukan transaksi debet) |
| `Kredit` | `#,##0.00` | Nilai kredit (0 jika bukan transaksi kredit) |
| `Saldo` | `#,##0.00` | Saldo setelah transaksi |

#### Cara kerja ekstraksi PDF

1. Membaca semua kata (`extract_words`) beserta koordinat X dan Y dari setiap halaman.
2. Mendeteksi posisi kolom `MUTASI` dan `SALDO` berdasarkan header di halaman pertama.
3. Mengelompokkan kata berdasarkan posisi Y → membentuk baris teks.
4. Mengidentifikasi awal transaksi baru berdasarkan pola tanggal `DD/MM` di awal baris.
5. Memisahkan teks ke kolom berdasarkan rentang koordinat X:
   - `x < x_mutasi_limit` → keterangan
   - `x_mutasi_limit ≤ x < x_saldo_limit` → nilai mutasi (Debet/Kredit)
   - `x ≥ x_saldo_limit` → saldo
6. Mendeteksi `DB`/`CR` dalam teks mutasi untuk menentukan kolom yang diisi.

> ⚠️ **Tahun hardcoded:** Kolom Tanggal menggunakan format `DD/MM/2025`. Jika memproses PDF dari tahun lain, ubah baris ini di `bca_extract_pdf.py`:
> ```python
> "Tanggal": date_match.group(1) + "/2025",
> # Ubah menjadi:
> "Tanggal": date_match.group(1) + "/2026",  # atau tahun yang sesuai
> ```

---

### 4. `gabung_BCA.py` — Gabungkan File Excel

Menggabungkan **semua file `.xlsx`** di folder menjadi satu workbook `Hasil_Gabungan_Mutasi_BCA.xlsx`, dengan satu sheet per kelompok rekening, diurutkan berdasarkan tanggal, dan saldo dihitung ulang secara kumulatif.

#### Cara pakai
1. Letakkan `gabung_BCA.py` di folder yang berisi semua file Excel mutasi BCA (hasil skrip 1, 2, atau 3).
2. Jalankan skrip.

#### Pengelompokan sheet

File dikelompokkan berdasarkan **angka pertama yang ditemukan di nama file**:

| Nama File | Sheet yang dihasilkan |
|---|---|
| `BCA 3456 15 AGU.xlsx` | `3456` |
| `BCA 3456 16 AGU.xlsx` | `3456` (digabung ke sheet yang sama) |
| `BCA 7890 15 AGU.xlsx` | `7890` |

Jika tidak ada angka di nama file, masuk ke sheet `Lainnya`.

#### Deteksi header otomatis

Skrip mencari baris header yang mengandung `Tanggal Transaksi` **dan** `Keterangan` dalam 25 baris pertama file Excel — tidak tergantung pada nomor baris yang tetap.

#### Penanganan kolom Jumlah (format DB/CR)

Jika file Excel memiliki kolom `Jumlah` berisi nilai seperti `1.234,56 CR` atau `1.234,56 DB`, nilai tersebut dipecah menjadi:
- Nilai **CR** → kolom `Kredit` (positif)
- Nilai **DB** → kolom `Debit` (absolut)

#### Rekalkuasi saldo

Setelah penggabungan dan pengurutan, saldo dihitung ulang dari awal secara kumulatif:
```
Saldo Awal = Saldo baris pertama − (Kredit₁ − Debit₁)
Saldo tiap baris = Saldo Awal + cumsum(Kredit − Debit)
```

#### Format output

- Kolom `Debit`, `Kredit`, `Saldo` → format angka `#,##0.00`
- Kolom `Tanggal Transaksi` → teks `DD/MM/YYYY`
- Auto-fit lebar semua kolom

---

### 5. `cek_tarikan_BCA2BCA.py` — Deteksi Transfer Antar Rekening

Menganalisis `Hasil_Gabungan_Mutasi_BCA.xlsx` untuk mendeteksi **transaksi transfer antar rekening BCA** yang muncul sebagai debit di satu rekening dan kredit di rekening lain pada tanggal dan nominal yang sama.

#### Cara pakai
1. Pastikan `Hasil_Gabungan_Mutasi_BCA.xlsx` ada di folder yang sama (hasil dari `gabung_BCA.py`).
2. Jalankan `cek_tarikan_BCA2BCA.py`.

#### Algoritma pencocokan

```
1. Pisahkan semua baris menjadi dua grup:
   - df_debit  → baris dengan Debit > 0
   - df_kredit → baris dengan Kredit > 0

2. JOIN pada (Tanggal Transaksi, Nominal) yang sama

3. Buang pasangan dari rekening yang sama
   (Bank_Pengirim == Bank_Penerima → bukan transfer antar rekening)

4. Untuk setiap pasangan yang tersisa:
   Hitung SequenceMatcher ratio antara Ket_Pengirim dan Ket_Penerima
   Jika similarity ≥ 0.70 → konfirmasi sebagai transfer valid
```

#### Kolom output (`Laporan_Transfer_Antar_Bank.xlsx`)

| Kolom | Keterangan |
|---|---|
| `Tanggal Transaksi` | Tanggal transfer |
| `Bank_Pengirim` | Nama sheet (kode rekening) asal debit |
| `Ket_Pengirim` | Keterangan dari sisi pengirim |
| `Nominal` | Nilai transfer |
| `Bank_Penerima` | Nama sheet (kode rekening) asal kredit |
| `Ket_Penerima` | Keterangan dari sisi penerima |

#### Pratinjau di terminal

Setiap transaksi yang terdeteksi ditampilkan di terminal sebelum disimpan ke Excel:

```
Tanggal    : 15/08/2025
Nominal    : 5,000,000.00
Dari Bank  : 3456
Ke Bank    : 7890
Ket (Kirim): TRF KE 7890 SAFFIELA
Ket (Trm)  : TRF DARI 3456
------------------------------------------------------------
```

---

## 🗺️ Alur Kerja yang Disarankan

### Jalur A — Dari CSV

```
1. Letakkan semua *.csv BCA di satu folder
2. Jalankan bcacsv2excel.py  (atau bcacsv2exceldbcr.py jika perlu DB/CR terpisah)
3. Hasilkan *.xlsx per CSV
4. Jalankan gabung_BCA.py → Hasil_Gabungan_Mutasi_BCA.xlsx
5. Jalankan cek_tarikan_BCA2BCA.py → Laporan_Transfer_Antar_Bank.xlsx
```

### Jalur B — Dari PDF

```
1. Letakkan semua *.pdf BCA di satu folder
2. Jalankan bca_extract_pdf.py → pilih mode (per dokumen atau semua)
3. Hasilkan {nama}_Excel.xlsx per PDF
4. Jalankan gabung_BCA.py → Hasil_Gabungan_Mutasi_BCA.xlsx
5. Jalankan cek_tarikan_BCA2BCA.py → Laporan_Transfer_Antar_Bank.xlsx
```

### Jalur C — Campuran CSV + PDF

```
1. Konversi CSV dengan bcacsv2excel.py atau bcacsv2exceldbcr.py
2. Konversi PDF dengan bca_extract_pdf.py
3. Kumpulkan semua *.xlsx hasil konversi ke satu folder
4. Jalankan gabung_BCA.py → Hasil_Gabungan_Mutasi_BCA.xlsx
5. Jalankan cek_tarikan_BCA2BCA.py → Laporan_Transfer_Antar_Bank.xlsx
```

---

## 🛠️ Troubleshooting

### ❌ `Tidak ditemukan file CSV di folder ini`
Pastikan file CSV ada di **folder yang sama** dengan skrip Python, bukan di subfolder.

### ❌ Nama output `BCA UNK 00 UNK.xlsx`
Skrip gagal membaca nomor rekening (baris ke-2) atau periode (baris ke-4) dari CSV. Buka file CSV dan periksa apakah format baris tersebut sesuai ekspor standar BCA.

### ❌ `Tidak ada PDF ditemukan`
Letakkan file PDF di folder yang sama dengan `bca_extract_pdf.py`. Ekstensi harus `.pdf` (huruf kecil).

### ❌ Kolom Tanggal di hasil PDF berisi tahun yang salah
Tahun 2025 hardcoded di `bca_extract_pdf.py`. Ubah baris:
```python
"Tanggal": date_match.group(1) + "/2025",
```
sesuai tahun dokumen PDF Anda.

### ❌ `File Hasil_Gabungan_Mutasi_BCA.xlsx tidak ditemukan` (di `cek_tarikan_BCA2BCA.py`)
Jalankan `gabung_BCA.py` terlebih dahulu. File ini adalah input wajib untuk deteksi transfer.

### ❌ Data hilang saat `gabung_BCA.py` — sheet kosong atau tidak ada file valid
Skrip mencari header `Tanggal Transaksi` dan `Keterangan` dalam 25 baris pertama. Jika file Excel hasil konversi tidak memiliki kedua header tersebut, file tersebut akan dilewati. Gunakan `gabung_BCA.py` hanya dengan file hasil dari skrip-skrip dalam proyek ini.

### ❌ `Tidak ditemukan transaksi transfer antar akun yang cocok`
Kemungkinan: (1) threshold kemiripan teks (0.7) terlalu tinggi untuk format keterangan di rekening Anda; (2) keterangan antar rekening terlalu berbeda. Turunkan nilai `0.7` di `cek_tarikan_BCA2BCA.py`:
```python
if similarity >= 0.7:  # Coba turunkan ke 0.5 atau 0.4
```

### ❌ `ModuleNotFoundError: No module named 'pdfplumber'`
```bash
pip install pdfplumber
```

---

## 📌 Catatan Penting

- **Setiap skrip berdiri sendiri** — tidak ada dependensi antar skrip kecuali `cek_tarikan_BCA2BCA.py` yang membutuhkan output dari `gabung_BCA.py`.
- **`gabung_BCA.py` mengecualikan dirinya sendiri** — file `Hasil_Gabungan_Mutasi_BCA.xlsx` dan file yang namanya diawali `~$` (file sementara Excel) tidak akan ikut diproses.
- **Tahun PDF hardcoded** — `bca_extract_pdf.py` selalu menulis `/2025` sebagai tahun di kolom Tanggal. Ubah sesuai tahun dokumen sebelum menjalankan.
- **Format CSV BCA bisa bervariasi** — Jika BCA mengubah format ekspor CSV-nya, posisi baris nomor rekening atau periode mungkin bergeser. Sesuaikan indeks baris di `bcacsv2excel.py` jika penamaan output tidak akurat.
- **Deteksi transfer bergantung pada kemiripan teks** — `cek_tarikan_BCA2BCA.py` menggunakan `difflib.SequenceMatcher` dengan threshold 70%. Transaksi dengan keterangan yang sangat berbeda antar rekening tidak akan terdeteksi meski nominal dan tanggalnya sama.
- **Rekalkuasi saldo di `gabung_BCA.py`** — Saldo dihitung ulang dari awal berdasarkan saldo baris pertama. Jika baris pertama tidak mewakili awal periode, saldo kumulatif akan meleset.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

* Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)
