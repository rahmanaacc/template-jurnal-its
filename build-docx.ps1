# =====================================================================
#  build-docx.ps1 — LaTeX -> DOCX (Pandoc) untuk Prosiding gaya ITS.
#  Jalankan:  ./build-docx.ps1   (atau VSCode: Run Task "LaTeX to DOCX")
#  Alur: Pandoc (sitasi IEEE + gambar + persamaan OMML, styling reference.docx)
#        lalu pasca-proses layout 2 kolom (tools/layout_2col.py).
# =====================================================================
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

# --- cari Pandoc & Python dari PATH (portabel antar-komputer) ---
$pandocCmd = Get-Command pandoc -ErrorAction SilentlyContinue
if (-not $pandocCmd) {
    Write-Host "Pandoc tidak ditemukan. Pasang dulu:" -ForegroundColor Red
    Write-Host "  winget install --id JohnMacFarlane.Pandoc" -ForegroundColor Yellow
    exit 1
}
$pandoc = $pandocCmd.Source

# Utamakan launcher `py` (cara standar Windows -> CPython resmi), lalu `python`.
# Pilih interpreter yang BENAR-BENAR punya modul python-docx.
$py = $null
foreach ($cand in @("py", "python", "python3")) {
    $c = Get-Command $cand -ErrorAction SilentlyContinue
    if (-not $c) { continue }
    & $c.Source -c "import docx" 2>$null
    if ($LASTEXITCODE -eq 0) { $py = $c.Source; break }
}
if (-not $py) {
    Write-Host "Python dengan modul 'python-docx' tidak ditemukan. Pasang dulu:" -ForegroundColor Red
    Write-Host "  py -m pip install python-docx" -ForegroundColor Yellow
    exit 1
}

# TODO: ganti nama berkas output sesuai kebutuhan (mis. NRP_Nama_Prosiding.docx)
$out = "Prosiding.docx"

Write-Host "[1/3] Pandoc: main.tex -> $out" -ForegroundColor Cyan
& $pandoc main.tex -o $out `
    --reference-doc=reference.docx `
    --citeproc --bibliography=pustaka/pustaka.bib --csl=ieee.csl `
    -M reference-section-title="DAFTAR PUSTAKA" `
    -M abstract-title="" `
    --resource-path=".;gambar"

Write-Host "[2/3] Pasca-proses layout 2 kolom" -ForegroundColor Cyan
& $py "tools\layout_2col.py" $out

Write-Host "[3/3] Preview PDF + jumlah halaman (butuh MS Word)" -ForegroundColor Cyan
try {
    $w = New-Object -ComObject Word.Application; $w.Visible = $false
    $d = $w.Documents.Open((Join-Path $PSScriptRoot $out))
    $pages = $d.ComputeStatistics(2)
    $d.SaveAs([ref](Join-Path $PSScriptRoot "preview.pdf"), [ref]17)
    $d.Close($false); $w.Quit()
    Write-Host "Selesai. Halaman: $pages (template Jurnal Teknik ITS: 6-8 hal A4)  ->  $out" -ForegroundColor Green
} catch {
    Write-Host "Docx dibuat: $out (Word tidak tersedia untuk preview)" -ForegroundColor Yellow
}
