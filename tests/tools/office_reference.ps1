# Dev-only: produce reference PDFs with the installed Microsoft Office so the
# pure-Python engine can be compared against it.  Not used by the library.
#   powershell -File tests/tools/office_reference.ps1 tests/fixtures
param([string]$Dir = "tests/fixtures")
$Dir = (Resolve-Path $Dir).Path
$ref = Join-Path $Dir "reference"
New-Item -ItemType Directory -Force $ref | Out-Null

$docs = Get-ChildItem $Dir -Filter *.docx | Where-Object { $_.Name -notlike "~$*" }
if ($docs) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($f in $docs) {
        $out = [string](Join-Path $ref ($f.BaseName + ".pdf"))
        $d = $word.Documents.Open([string]$f.FullName, $false, $true, $false)
        $d.ExportAsFixedFormat($out, 17)
        $d.Close(0)
        Write-Output "word: $out"
    }
    $word.Quit()
}

$books = Get-ChildItem $Dir -Filter *.xlsx | Where-Object { $_.Name -notlike "~$*" }
if ($books) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    foreach ($f in $books) {
        $out = [string](Join-Path $ref ($f.BaseName + ".pdf"))
        $wb = $xl.Workbooks.Open([string]$f.FullName)
        # Save once so formula results are cached in the file for the engine.
        $wb.Save()
        $wb.ExportAsFixedFormat(0, $out)
        $wb.Close($false)
        Write-Output "excel: $out"
    }
    $xl.Quit()
}

