# Folder gambar

Letakkan semua gambar naskah di sini, lalu panggil dari `.tex` dengan:

```latex
\includegraphics[width=8.5cm]{gambar/nama-file.png}
```

Catatan:
- Gunakan format **.png** atau **.jpg** (Pandoc melewati `.pdf`/`.eps`).
- Lebar **8,5 cm** pas untuk satu kolom; gunakan lebih kecil bila perlu.
- `contoh-gambar.png` hanya placeholder — ganti dengan gambar Anda dan hapus
  yang tidak terpakai.

## Gambar vektor (SVG)

Untuk diagram dan ilustrasi vektor, letakkan file **.svg** di folder ini.
Build script akan **otomatis mengonversi** `.svg` → `.png` menggunakan
`rsvg-convert` sebelum Pandoc dijalankan. Anda tetap memanggil file `.png`
di LaTeX:

```latex
% diagram.svg (sumber) -> otomatis jadi diagram.png saat build
\includegraphics[width=8.5cm]{gambar/diagram.png}
```

> **Syarat**: `rsvg-convert` harus tersedia (MSYS2: `pacman -S mingw-w64-x86_64-librsvg`).
> Jika tidak ada file `.svg`, langkah ini dilewati otomatis.
