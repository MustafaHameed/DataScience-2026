# Build the Advanced Research Methodology master handout.
#   .\build.ps1          normal build
#   .\build.ps1 -Clean   remove aux files first
param([switch]$Clean)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$main = 'AdvancedResearchMethodology_Handout'

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

# latexmk narrates progress on stderr ("Missing input file ... .toc" on the
# first pass of a clean build, and similar). Windows PowerShell turns any
# stderr line from a native exe into an ErrorRecord, which $ErrorActionPreference
# = 'Stop' then treats as fatal. Relax the preference around the native calls
# only, and judge success by the exit code instead.
function Invoke-Latexmk {
    param([string[]]$Arguments)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $errFile = [System.IO.Path]::GetTempFileName()
    try {
        # Out-Host, not bare output: the caller wants the exit code back, not
        # latexmk's several hundred lines of chatter. stderr goes to a file so
        # that routine progress notes ("Missing input file ... .toc" on the
        # first pass of a clean build) are not rendered as red error blocks;
        # it is replayed in full if the build actually fails.
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

Write-Host "Building $main.pdf ..." -ForegroundColor Cyan
$latexmkExit = Invoke-Latexmk @('-pdf', '-interaction=nonstopmode', '-file-line-error', $main)

$log = "$main.log"
if ($latexmkExit -ne 0 -and -not (Test-Path $log)) {
    Write-Host "latexmk failed (exit $latexmkExit) and wrote no log." -ForegroundColor Red
    exit $latexmkExit
}
if (-not (Test-Path $log)) {
    Write-Host 'No log file was produced; nothing to report.' -ForegroundColor Red
    exit 1
}

# A stale log from an earlier run would make a failed build look clean, so make
# the report refuse to speak for a log it did not just produce.
$logAge = (Get-Date) - (Get-Item $log).LastWriteTime
if ($logAge.TotalMinutes -gt 30) {
    Write-Host "Log is $([int]$logAge.TotalMinutes) minutes old - this build produced nothing." -ForegroundColor Red
    exit 1
}

# "Infinite glue shrinkage found in box being split" is upstream longtable
# behaviour whenever a table spans a page break: TeX drops the shrink component
# and the rules and repeated headers still come out correct. Filter it out so
# that real errors stay visible.
$errors = Select-String -Path $log -Pattern '^!|^.*:\d+: ' -ErrorAction SilentlyContinue |
          Where-Object { $_.Line -notmatch 'Infinite glue shrinkage' }
$undef     = Select-String -Path $log -Pattern 'undefined references|Reference .* undefined' -ErrorAction SilentlyContinue
$overH     = (Select-String -Path $log -Pattern 'Overfull \\hbox'  -ErrorAction SilentlyContinue | Measure-Object).Count
$overV     = (Select-String -Path $log -Pattern 'Overfull \\vbox'  -ErrorAction SilentlyContinue | Measure-Object).Count
$underH    = (Select-String -Path $log -Pattern '^Underfull'       -ErrorAction SilentlyContinue | Measure-Object).Count
$fontWarn  = (Select-String -Path $log -Pattern 'LaTeX Font Warning' -ErrorAction SilentlyContinue | Measure-Object).Count
$pdfString = (Select-String -Path $log -Pattern 'Token not allowed' -ErrorAction SilentlyContinue | Measure-Object).Count
$pagesHit  = Select-String -Path $log -Pattern 'Output written on .*\((\d+) pages' -ErrorAction SilentlyContinue |
             Select-Object -Last 1
$pages     = if ($pagesHit) { $pagesHit.Matches.Groups[1].Value } else { '(none - build produced no PDF)' }

$errCount = ($errors | Measure-Object).Count
$refCount = ($undef  | Measure-Object).Count

function Write-Metric($label, $value) {
    $colour = if ($value -is [int] -and $value -gt 0) { 'Yellow' } else { 'Gray' }
    Write-Host ("{0,-24}{1}" -f $label, $value) -ForegroundColor $colour
}

Write-Host ''
Write-Host '--- Build report ---' -ForegroundColor Cyan
Write-Host ("{0,-24}{1}" -f 'Pages:', $pages)
Write-Metric 'Errors:'            $errCount
Write-Metric 'Undefined refs:'    $refCount
Write-Metric 'Overfull hboxes:'   $overH
Write-Metric 'Overfull vboxes:'   $overV
Write-Metric 'Underfull boxes:'   $underH
Write-Metric 'Font warnings:'     $fontWarn
Write-Metric 'PDF-string warns:'  $pdfString

if ($errCount -gt 0) {
    Write-Host ''
    Write-Host 'First errors:' -ForegroundColor Red
    $errors | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.Line)" }
    exit 1
}

if ($latexmkExit -ne 0) { exit $latexmkExit }
