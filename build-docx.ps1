# =====================================================================
#  build-docx.ps1 — LaTeX -> DOCX (Pandoc) untuk Prosiding gaya ITS.
#  Jalankan:  ./build-docx.ps1   (atau VSCode: Run Task "LaTeX to DOCX")
#  Alur: SVG->PNG -> Pandoc (sitasi IEEE + gambar + persamaan OMML)
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

# --- cek rsvg-convert untuk SVG->PNG ---
$rsvg = "C:\msys64\mingw64\bin\rsvg-convert.exe"
if (-not (Test-Path -LiteralPath $rsvg)) {
    Write-Host "rsvg-convert tidak ditemukan di $rsvg" -ForegroundColor Red
    Write-Host "Pasang: pacman -S mingw-w64-x86_64-librsvg (MSYS2)" -ForegroundColor Yellow
    exit 1
}

# --- 0) Konversi SVG -> PNG (sumber vektor, output raster untuk Pandoc) ---
Write-Host "[0/3] SVG -> PNG (rsvg-convert)" -ForegroundColor Cyan
Get-ChildItem -LiteralPath "gambar" -Filter "*.svg" | ForEach-Object {
    $pngPath = Join-Path $_.Directory.FullName ($_.BaseName + ".png")
    Write-Host "  $($_.Name) -> $($_.BaseName).png"
    & $rsvg -w 1200 $_.FullName -o $pngPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  GAGAL: $($_.Name)" -ForegroundColor Red
        exit 1
    }
}
Write-Host ""

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
    $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.ScreenUpdating = $false
    $d = $w.Documents.Open((Join-Path $PSScriptRoot $out))
    $pages = $d.ComputeStatistics(2)
    $d.SaveAs([ref](Join-Path $PSScriptRoot "preview.pdf"), [ref]17)
    $d.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($d) | Out-Null
    $w.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($w) | Out-Null
    Write-Host "Selesai. Halaman: $pages (template Jurnal Teknik ITS: 6-8 hal A4)  ->  $out" -ForegroundColor Green
} catch {
    Write-Host "Docx dibuat: $out (Word tidak tersedia untuk preview)" -ForegroundColor Yellow
}
