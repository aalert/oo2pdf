# Dev-only: build a local corpus from the templates shipped with Office and
# export Office's own PDFs as references.  The corpus stays local (it is
# Microsoft content); tests/corpus is git-ignored.
#   powershell -File tests/tools/corpus_reference.ps1
param([string]$Out = "tests/corpus")
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = [string](Resolve-Path $Out).Path
$ref = [string](Join-Path $Out "reference")
New-Item -ItemType Directory -Force $ref | Out-Null
$src = Get-ChildItem "C:\Program Files\Microsoft Office\root\Templates" -Recurse -Include *.dotx, *.xltx

$words = $src | Where-Object { $_.Extension -ieq ".dotx" }
if ($words) {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    foreach ($f in $words) {
        $name = $f.BaseName -replace '\s', '_'
        $docx = [string](Join-Path $Out ($name + ".docx"))
        $pdf = [string](Join-Path $ref ($name + ".pdf"))
        try {
            $d = $w.Documents.Add([string]$f.FullName)
            $d.SaveAs2($docx, 16)
            $d.ExportAsFixedFormat($pdf, 17)
            $d.Close(0)
            Write-Output "word ok: $name"
        } catch { Write-Output "word FAILED: $name $_" }
    }
    $w.Quit()
}

$books = $src | Where-Object { $_.Extension -ieq ".xltx" }
if ($books) {
    $x = New-Object -ComObject Excel.Application
    $x.Visible = $false
    $x.DisplayAlerts = $false
    foreach ($f in $books) {
        $name = $f.BaseName -replace '\s', '_'
        $xlsx = [string](Join-Path $Out ($name + ".xlsx"))
        $pdf = [string](Join-Path $ref ($name + "_xlsx.pdf"))
        try {
            $b = $x.Workbooks.Add([string]$f.FullName)
            $b.SaveAs($xlsx, 51)
            $b.ExportAsFixedFormat(0, $pdf)
            $b.Close($false)
            Write-Output "excel ok: $name"
        } catch { Write-Output "excel FAILED: $name $_" }
    }
    $x.Quit()
}
