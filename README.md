# Template Prosiding / Jurnal Teknik ITS (LaTeX → DOCX)

Template untuk menulis artikel **Jurnal Teknik ITS** gaya POMITS.
Naskah ditulis dalam **LaTeX**, lalu dikonversi otomatis ke **Word (.docx)**
dengan Pandoc.

> Kenapa LaTeX → DOCX? Banyak jurnal/prosiding mewajibkan submit **.docx**,
> tetapi menulis di Word merepotkan untuk persamaan dan sitasi. Template ini
> memberi kemudahan menulis LaTeX dengan hasil akhir `.docx` yang sudah sesuai
> format.

## Prasyarat (sekali pasang)

- **Pandoc** → `winget install --id JohnMacFarlane.Pandoc`
- **Python** + **python-docx** → `py -m pip install python-docx`
- **Microsoft Word** *(opsional)* — hanya untuk preview PDF & hitung halaman.

## Cara pakai (3 langkah)

1. **Edit identitas** di `main.tex` (judul, penulis, afiliasi, e-mail).
2. **Tulis naskah** di `bagian/*.tex`.
3. **Build:**
   ```powershell
   ./build-docx.ps1
   ```
   Hasil: **`Prosiding.docx`** (+ `preview.pdf` bila ada Word).

> Cari penanda **`TODO`** di seluruh proyek untuk menemukan semua tempat yang
> perlu Anda ganti.

## Struktur folder

```text
├── main.tex                 # dokumen utama: judul, penulis, \input tiap bagian
├── bagian/                  # ISI NASKAH — edit di sini
│   ├── 00-abstrak.tex
│   ├── 01-pendahuluan.tex
│   ├── 02-metodologi.tex
│   ├── 03-hasil-pembahasan.tex
│   └── 04-kesimpulan.tex
├── gambar/                  # gambar (.png/.jpg) + contoh placeholder
├── pustaka/pustaka.bib      # basis referensi (BibTeX)
├── reference.docx           # ⚙ template gaya ITS — JANGAN dihapus
├── ieee.csl                 # gaya sitasi IEEE
├── build-docx.ps1           # 1-klik: Pandoc → layout 2 kolom → preview PDF
├── tools/
│   ├── layout_2col.py       # pasca-proses 2 kolom, header, drop cap, dll.
│   └── style_reference.py   # (lanjutan) membangun ulang reference.docx
└── .vscode/                 # Run Task "LaTeX to DOCX"
```

## Konvensi gaya ITS

- Penomoran bab **Romawi manual + HURUF KAPITAL**: `\section{I. PENDAHULUAN}`
- Subbab huruf: `\subsection{A. ...}`
- Nomor Gambar/Tabel/Persamaan **ditulis manual** pada caption (bukan `\label/\ref`)
- Huruf pertama Pendahuluan otomatis dijadikan **drop cap**

## Lisensi

Bebas dipakai dan dimodifikasi untuk keperluan akademik. Lihat `LICENSE`.
