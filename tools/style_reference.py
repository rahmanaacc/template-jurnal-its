# -*- coding: utf-8 -*-
"""Restyle reference.docx Pandoc -> gaya Jurnal Teknik ITS (POMITS).
Mengatur named-style yang dipakai Pandoc: Title, Author, Abstract,
Heading 1/2, Body Text/First Paragraph, caption, Bibliography, dll.
Halaman A4 + margin ITS. (Kolom 2 untuk isi diatur oleh layout_2col.py.)

PERINGATAN: `reference.docx` SUDAH disetel-tangan melampaui skrip ini
(mis. Body Text/Abstract spacing-after=0, caption 9pt rata-kiri-kanan,
spacing heading khusus) -- nilai-nilai itu yang menjaga drop cap "M" pas
3 baris. MENJALANKAN ULANG skrip ini akan MENGEMBALIKAN tweak tersebut
(after=120, caption 8pt center) dan merusak layout. Jangan jalankan lagi
kecuali ingin membangun ulang reference.docx dari nol. Perbaikan warna
teks (heading/hyperlink -> hitam) sekarang ditangani layout_2col.py
setiap build, jadi tidak perlu skrip ini.
"""
import sys, os, re, zipfile
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL

PATH = sys.argv[1]
TNR = "Times New Roman"


def patch_theme_fonts(path, font=TNR):
    """Set major & minor (Latin) theme font -> TNR.

    Style Pandoc memakai rFonts dengan w:asciiTheme="majorHAnsi"/"minorHAnsi".
    Bila atribut theme ADA, Word memprioritaskannya di atas w:ascii eksplisit,
    sehingga judul/heading ikut font theme (default Word = 'Aptos Display'),
    BUKAN Times New Roman. Mengganti font theme -> TNR membuat semua teks yang
    mereferensikan theme tetap Times New Roman.
    """
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    tkey = 'word/theme/theme1.xml'
    if tkey in data:
        t = data[tkey].decode('utf-8')
        for block in ('majorFont', 'minorFont'):
            t = re.sub(r'(<a:' + block + r'>.*?<a:latin typeface=")[^"]*(")',
                       r'\g<1>' + font + r'\g<2>', t, flags=re.S)
        data[tkey] = t.encode('utf-8')
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    os.replace(tmp, path)
    print("Theme major/minor font ->", font)

doc = Document(PATH)
names = {s.name for s in doc.styles}

def style(name, size=None, bold=None, italic=None, align=None, caps=False,
          after=None, before=None, line=1.0, hanging=None):
    if name not in names:
        return
    st = doc.styles[name]
    f = st.font
    f.name = TNR
    if size is not None: f.size = Pt(size)
    if bold is not None: f.bold = bold
    if italic is not None: f.italic = italic
    rpr = st.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), TNR)
    if caps:
        c = OxmlElement('w:caps'); c.set(qn('w:val'), 'true'); rpr.append(c)
    pf = st.paragraph_format
    if align is not None: pf.alignment = align
    if after is not None: pf.space_after = Pt(after)
    if before is not None: pf.space_before = Pt(before)
    if line is not None: pf.line_spacing = line
    if hanging is not None:
        pf.left_indent = Cm(hanging); pf.first_line_indent = Cm(-hanging)

# Base & body
style("Normal",         size=10, align=AL.JUSTIFY, after=0, line=1.0)
style("Body Text",      size=10, align=AL.JUSTIFY, after=6, line=1.0)
style("First Paragraph",size=10, align=AL.JUSTIFY, after=0, line=1.0)
style("Compact",        size=10, align=AL.JUSTIFY, after=0, line=1.0)
# Front matter
style("Title",   size=24, bold=False, align=AL.CENTER, after=6, line=1.0)
style("Author",  size=10, align=AL.CENTER, after=0, line=1.0)
style("Abstract",size=9,  bold=True, align=AL.JUSTIFY, after=6, line=1.0)
# Headings
style("Heading 1", size=10, bold=True,  italic=False, align=AL.CENTER,
      caps=True, before=8, after=3, line=1.0)
style("Heading 2", size=10, bold=False, italic=True,  align=AL.LEFT,
      before=4, after=2, line=1.0)
style("Heading 3", size=10, bold=False, italic=True,  align=AL.LEFT,
      before=3, after=2, line=1.0)
# Captions (gambar/tabel) & referensi
for cap in ("Caption", "Image Caption", "Table Caption"):
    style(cap, size=8, italic=False, align=AL.CENTER, before=2, after=6, line=1.0)
style("Bibliography", size=8, align=AL.JUSTIFY, after=2, line=1.0, hanging=0.6)

# Halaman A4 + margin ITS untuk semua section
for sec in doc.sections:
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.65)
    sec.top_margin = sec.bottom_margin = Cm(1.78)

doc.save(PATH)
patch_theme_fonts(PATH)          # judul/heading -> TNR (bukan font theme)
print("reference.docx distyle ITS:", PATH)
print("style tersedia:", sorted(n for n in names if n in {
    'Title','Author','Abstract','Heading 1','Heading 2','Body Text',
    'Image Caption','Table Caption','Caption','Bibliography','Normal'}))
