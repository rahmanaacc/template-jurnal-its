# Template Jurnal Teknik ITS — Memory & Context

## Prasyarat
- Pandoc (`winget install --id JohnMacFarlane.Pandoc`)
- Python + `python-docx` (`py -m pip install python-docx`)
- Microsoft Word (untuk preview PDF + halaman)
- **rsvg-convert** — untuk konversi SVG→PNG (via MSYS2: `pacman -S mingw-w64-x86_64-librsvg`). Opsional: hanya jika menggunakan file `.svg`.

## Build
```powershell
./build-docx.ps1
```
Output: `Prosiding.docx` + `preview.pdf` (6–8 halaman)

---

## Struktur Proyek
```
├── main.tex                         # Entry point LaTeX
├── bagian/
│   ├── 00-abstrak.tex               # Abstrak + Kata Kunci (bold)
│   ├── 01-pendahuluan.tex           # I. PENDAHULUAN
│   ├── 02-metodologi.tex            # II. URAIAN PENELITIAN
│   ├── 03-hasil-pembahasan.tex      # III. HASIL DAN DISKUSI
│   └── 04-kesimpulan.tex            # IV. KESIMPULAN
├── gambar/
│   ├── *.svg                        # Sumber vektor — dikonversi otomatis ke PNG saat build
│   ├── *.png                        # Output raster & gambar non-vektor
│   └── contoh-gambar.png            # Placeholder — ganti dengan gambar Anda
├── tools/
│   ├── layout_2col.py               # Pasca-proses DOCX: 2 kolom, header ITS, dll
│   └── style_reference.py           # Generator reference.docx (jangan dijalankan ulang)
├── pustaka/pustaka.bib              # BibTeX references
├── reference.docx                   # Template styling DOCX (jangan diedit)
├── ieee.csl                         # Citation style IEEE
└── build-docx.ps1                   # Script build utama
```

---

## Konvensi Penulisan (Jurnal Teknik ITS)
- Penomoran bab: ROMAWI manual (`\section{I. PENDAHULUAN}`)
- Penomoran subbab: HURUF (`\subsection{A. Arsitektur Sistem}`)
- Penomoran gambar/tabel: manual dalam caption (bukan `\label`/`\ref`)
- Caption gambar: 8pt non-bold (diatur di `layout_2col.py`)
- Heading warna hitam (000000), bukan biru theme (diatur di `layout_2col.py`)
- Abstrak + Kata Kunci: bold (diatur di `layout_2col.py`)
- Alinea badan teks: first-line indent 10,1pt
- Drop cap: huruf pertama Pendahuluan otomatis 3 baris
- Border tabel: horizontal *booktabs* via `layout_2col.py` (atas, bawah header, bawah)
- Caption tabel: rata tengah

## Konvensi Penulisan Ilmiah
- **Perincian**: `1. 2. 3.` untuk urutan/prosedur, `a. b. c.` untuk item setara. JANGAN bullet (• - *).
- **Bold HANYA untuk**: judul bernomor, caption, penegasan istilah/variabel, akronim baru.
- **Bold DILARANG**: menekankan kata "penting", istilah asing belum diserap (pakai *italic*), seluruh kalimat/rumus.
- **Bahasa**: Indonesia baku (PUEBI/EYD V), koma desimal (3,14), singkatan didefinisikan saat pertama muncul.
- **Klaim kuantitatif** wajib didukung data/rujukan.

---

## Troubleshooting
- **Preview tidak update**: Tutup semua Word + PDF reader yang membuka `Prosiding.docx`/`preview.pdf`, lalu `taskkill /f /im WINWORD.EXE` bila perlu, baru jalankan `./build-docx.ps1`
