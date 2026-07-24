# Template Prosiding / Jurnal Teknik ITS (LaTeX → DOCX)

Template untuk menulis artikel **Prosiding Seminar Tugas Akhir / Jurnal Teknik
ITS** gaya POMITS. Naskah ditulis dalam **LaTeX** (rapi & *expandable*), lalu
dikonversi otomatis ke **Word (.docx)** dengan Pandoc. Styling (Times New Roman,
heading, A4, 2 kolom, header tiap halaman) sudah diatur — Anda cukup mengisi
teks dan menjalankan satu perintah.

> Kenapa LaTeX → DOCX? Banyak jurnal/prosiding mewajibkan submit **.docx**,
> tetapi menulis di Word merepotkan untuk persamaan dan sitasi. Template ini
> memberi kemudahan menulis LaTeX dengan hasil akhir berupa `.docx` yang sudah
> sesuai format.

## Hasil yang dijamin

| Syarat format | Status |
|---|---|
| Format **.docx** | ✅ |
| Ukuran **A4** (21 × 29,7 cm) | ✅ |
| Font **Times New Roman** | ✅ |
| Judul 24pt non-bold rata tengah (Title Case) | ✅ |
| 2 kolom + abstrak/judul *full-width* | ✅ |
| Heading bernomor (I., II., …) HURUF KAPITAL + subbab (A., B., …) | ✅ |
| Persamaan, subskrip, sitasi IEEE | ✅ otomatis dari LaTeX (OMML + citeproc) |
| Header Jurnal Teknik ITS + nomor halaman otomatis | ✅ |

> Sesuai *gaya selingkung* **Jurnal Teknik ITS** (`TEMPLATE_PUBLIKASI_TEKNIK`):
> panjang artikel **6–8 halaman A4**. Sebagian acara (mis. Seminar Tugas Akhir)
> menetapkan batas sendiri — periksa ketentuan acara Anda.

## Prasyarat (sekali pasang)

- **Pandoc** → `winget install --id JohnMacFarlane.Pandoc`
- **Python** + **python-docx** → `python -m pip install python-docx`
- **Microsoft Word** *(opsional)* — hanya untuk preview PDF & hitung halaman.

## Cara pakai (3 langkah)

1. **Edit identitas** di `main.tex` (judul, penulis, afiliasi, e-mail) dan
   header kop (Vol/No/tahun) di `tools/layout_2col.py` — variabel `HEADER_TEXT`
   pada blok `KONFIGURASI`. Header gaya **Jurnal Teknik ITS** (teks di kiri +
   nomor halaman otomatis di kanan).
2. **Tulis naskah** di `bagian/*.tex` (lihat panduan di bawah).
3. **Build:**
   ```powershell
   ./build-docx.ps1
   ```
   atau di VSCode: `Ctrl+Shift+P` → *Run Task* → **LaTeX to DOCX**.

   Hasil: **`Prosiding.docx`** (+ `preview.pdf` bila ada Word).

> Cari penanda **`TODO`** di seluruh proyek untuk menemukan semua tempat yang
> perlu Anda ganti.

## Struktur folder

```text
prosiding-its-template/
├─ main.tex                 # dokumen utama: judul, penulis, \input tiap bagian
├─ bagian/                  # ISI NASKAH — edit di sini (expandable)
│  ├─ 00-abstrak.tex
│  ├─ 01-pendahuluan.tex
│  ├─ 02-metodologi.tex     # contoh gambar, persamaan, tabel
│  ├─ 03-hasil-pembahasan.tex
│  └─ 04-kesimpulan.tex
├─ gambar/                  # gambar (.png/.jpg) + contoh placeholder
├─ pustaka/pustaka.bib      # basis referensi (BibTeX)
├─ reference.docx           # ⚙ template gaya ITS untuk Pandoc — JANGAN dihapus
├─ ieee.csl                 # gaya sitasi IEEE
├─ build-docx.ps1           # 1-klik: Pandoc → layout 2 kolom → preview PDF
├─ tools/
│  ├─ layout_2col.py        # pasca-proses 2 kolom, header, drop cap, dll.
│  └─ style_reference.py    # (lanjutan) membangun ulang reference.docx
└─ .vscode/                 # Run Task "LaTeX to DOCX"
```

## Panduan menulis (`bagian/*.tex`)

- **Teks biasa:** tulis LaTeX seperti umumnya.
- **Istilah asing (miring):** `\textit{Field-Oriented Control}`.
- **Persamaan:** `$...$` (inline) atau lingkungan `equation*` dengan nomor
  manual `\qquad\text{(1)}` — otomatis jadi *equation* Word.
- **Sitasi:** `\cite{kunci}` dengan kunci dari `pustaka/pustaka.bib`. Daftar
  Pustaka dibangkitkan otomatis (gaya IEEE), tak perlu ditulis tangan.
- **Gambar:** `\includegraphics[width=8.5cm]{gambar/nama.png}` di dalam
  lingkungan `figure`, dengan `\caption{Gambar N. ...}` (nomor manual).
- **Tabel:** lingkungan `table` + `tabular`; `\caption{Tabel N. ...}` di atas
  tabel. Akan dirapikan otomatis (rata-tengah, selebar kolom).

### Menambah bagian baru

1. Buat berkas `bagian/NN-nama.tex` (mis. `05-ucapan-terima-kasih.tex`).
2. Tambahkan `\input{bagian/NN-nama}` di `main.tex`.
3. Build ulang.

## Konvensi gaya ITS

- Penomoran bab **Romawi manual + HURUF KAPITAL** (sesuai template):
  `\section{I. PENDAHULUAN}`, `\section{II. METODOLOGI}`, dst.
- Subbab huruf: `\subsection{A. ...}`, `\subsection{B. ...}`.
- Nomor Gambar/Tabel/Persamaan **ditulis manual** pada caption (bukan
  `\label/\ref`), sesuai gaya template ITS.
- Huruf pertama Pendahuluan otomatis dijadikan **drop cap**.
- Slot **LAMPIRAN** otomatis ditambahkan di akhir (boleh dibiarkan/dihapus di
  Word).

## Bagaimana pipeline bekerja (singkat)

1. **Pandoc** mengubah `main.tex` → `.docx` (sitasi IEEE via `--citeproc` +
   `ieee.csl`, gambar dari `gambar/`, persamaan → OMML, styling
   `reference.docx`).
2. **`tools/layout_2col.py`** memasang *section break* (judul/abstrak 1 kolom,
   isi 2 kolom), meng-*center* gambar, merapikan tabel, memasang header & drop
   cap, lalu menambah slot Lampiran.
3. Ekspor `preview.pdf` + hitung halaman (bila MS Word tersedia).

> **Catatan tentang `reference.docx`:** berkas ini adalah "kerangka gaya" yang
> dipakai Pandoc dan sudah disetel-tangan. **Jangan dihapus.** `style_reference.py`
> hanya untuk membangunnya ulang dari nol (lanjutan) — lihat peringatan di
> dalam berkas itu sebelum menjalankannya.

## Lisensi

Bebas dipakai dan dimodifikasi untuk keperluan akademik. Lihat `LICENSE`.
