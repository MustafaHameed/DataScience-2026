# Build the Urdu brief edition of the AIOU 430 Principles of Journalism course guide.
#   .\build_urdu.ps1          normal build
#   .\build_urdu.ps1 -Clean   remove aux files first
#
# XeLaTeX, not pdfLaTeX: Urdu needs OpenType shaping (Nastaliq ligatures)
# and the bidirectional algorithm, neither of which pdfTeX has.
param([switch]$Clean)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$main = 'AIOU_430_Urdu_Brief'

# The brief editions are capped at 21 pages; the point of them is that they
# fit in one sitting. Enforce it rather than trusting a later eyeball.
$pageCap = 21

# latexmk is a Perl script. MiKTeX ships the script but not an interpreter, and
# a plain PowerShell session usually has no perl on PATH even when Git for
# Windows has one bundled. Borrow it rather than failing.
if (-not (Get-Command perl -ErrorAction SilentlyContinue)) {
    $candidates = @(
        "$env:ProgramFiles\Git\usr\bin",
        "${env:ProgramFiles(x86)}\Git\usr\bin",
        "$env:LOCALAPPDATA\Programs\Git\usr\bin",
        'C:\Strawberry\perl\bin'
    )
    $perlDir = $candidates | Where-Object { Test-Path (Join-Path $_ 'perl.exe') } | Select-Object -First 1
    if ($perlDir) {
        $env:PATH = "$perlDir;$env:PATH"
        Write-Host "Using perl from $perlDir" -ForegroundColor DarkGray
    } else {
        Write-Host 'No perl found. latexmk cannot run; install Strawberry Perl or Git for Windows.' -ForegroundColor Red
        exit 1
    }
}

# The document needs the "Urdu Typesetting" system font. Fail early and
# legibly rather than letting fontspec produce 400 lines of nullfont noise.
Add-Type -AssemblyName System.Drawing
$installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
if ($installed -notcontains 'Urdu Typesetting') {
    Write-Host 'Font "Urdu Typesetting" is not installed.' -ForegroundColor Red
    Write-Host 'It ships with Windows; on other systems substitute a Nastaliq or' -ForegroundColor Red
    Write-Host 'Urdu Naskh face in aiouurdu.sty (e.g. Noto Nastaliq Urdu).' -ForegroundColor Red
    exit 1
}

# latexmk narrates progress on stderr. Windows PowerShell turns any stderr line
# from a native exe into an ErrorRecord, which $ErrorActionPreference = 'Stop'
# then treats as fatal. Relax the preference around the native call only, and
# judge success by the exit code instead.
function Invoke-Latexmk {
    param([string[]]$Arguments)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $errFile = [System.IO.Path]::GetTempFileName()
    try {
        & latexmk @Arguments 2> $errFile | Out-Host
        $code = $LASTEXITCODE
        if ($code -ne 0 -and (Get-Item $errFile).Length -gt 0) {
            Get-Content $errFile | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        }
        return $code
    } finally {
        $ErrorActionPreference = $previous
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
    }
}

if ($Clean) {
    Invoke-Latexmk @('-C', $main) > $null
    Write-Host 'Cleaned auxiliary files.' -ForegroundColor Yellow
}

Write-Host "Building $main.pdf (XeLaTeX) ..." -ForegroundColor Cyan
$latexmkExit = Invoke-Latexmk @('-xelatex', '-interaction=nonstopmode', '-file-line-error', $main)

$log = "$main.log"
if (-not (Test-Path $log)) {
    Write-Host "latexmk failed (exit $latexmkExit) and wrote no log." -ForegroundColor Red
    exit 1
}

$logAge = (Get-Date) - (Get-Item $log).LastWriteTime
if ($logAge.TotalMinutes -gt 30) {
    Write-Host "Log is $([int]$logAge.TotalMinutes) minutes old - this build produced nothing." -ForegroundColor Red
    exit 1
}

$errors   = Select-String -Path $log -Pattern '^!|^.*:\d+: ' -ErrorAction SilentlyContinue
$undef    = Select-String -Path $log -Pattern 'undefined references|Reference .* undefined' -ErrorAction SilentlyContinue
$overH    = (Select-String -Path $log -Pattern 'Overfull \\hbox' -ErrorAction SilentlyContinue | Measure-Object).Count
$underH   = (Select-String -Path $log -Pattern '^Underfull'      -ErrorAction SilentlyContinue | Measure-Object).Count
# The signature failure mode for an Urdu build: the font did not load and every
# Arabic-script glyph silently vanished. Count it as a first-class metric.
$missing  = (Select-String -Path $log -Pattern 'Missing character' -ErrorAction SilentlyContinue | Measure-Object).Count
$fontWarn = (Select-String -Path $log -Pattern 'LaTeX Font Warning' -ErrorAction SilentlyContinue | Measure-Object).Count
$pagesHit = Select-String -Path $log -Pattern 'Output written on .*\((\d+) pages' -ErrorAction SilentlyContinue |
            Select-Object -Last 1
$pages    = if ($pagesHit) { $pagesHit.Matches.Groups[1].Value } else { '(none - build produced no PDF)' }

$errCount = ($errors | Measure-Object).Count
$refCount = ($undef  | Measure-Object).Count

function Write-Metric($label, $value) {
    $colour = if ($value -is [int] -and $value -gt 0) { 'Yellow' } else { 'Gray' }
    Write-Host ("{0,-24}{1}" -f $label, $value) -ForegroundColor $colour
}

Write-Host ''
Write-Host '--- Build report (430 Urdu brief) ---' -ForegroundColor Cyan
Write-Host ("{0,-24}{1}" -f 'Pages:', $pages)
Write-Metric 'Errors:'            $errCount
Write-Metric 'Undefined refs:'    $refCount
Write-Metric 'Missing glyphs:'    $missing
Write-Metric 'Overfull hboxes:'   $overH
Write-Metric 'Underfull boxes:'   $underH
Write-Metric 'Font warnings:'     $fontWarn

if ($missing -gt 0) {
    Write-Host ''
    Write-Host 'Missing glyphs means the Urdu font did not load - the text is dropped,' -ForegroundColor Red
    Write-Host 'not merely mis-shaped. Check the \setmainfont line in aiouurdu.sty.' -ForegroundColor Red
    exit 1
}

if ($errCount -gt 0) {
    Write-Host ''
    Write-Host 'First errors:' -ForegroundColor Red
    $errors | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.Line)" }
    exit 1
}

if ($pages -match '^\d+$' -and [int]$pages -gt $pageCap) {
    Write-Host ''
    Write-Host "Brief edition is $pages pages, over its $pageCap-page cap." -ForegroundColor Red
    Write-Host 'Cut a question or tighten an answer; do not raise the cap.' -ForegroundColor Red
    exit 1
}

if ($latexmkExit -ne 0) { exit $latexmkExit }
