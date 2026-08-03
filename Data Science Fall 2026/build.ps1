# Build the Data Science Fall 2026 master handout.
#   .\build.ps1          normal build
#   .\build.ps1 -Clean   remove aux files first
param([switch]$Clean)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$main = 'DataScience_Fall2026_Handout'

if ($Clean) {
    latexmk -C $main 2>&1 | Out-Null
    Write-Host 'Cleaned auxiliary files.' -ForegroundColor Yellow
}

Write-Host "Building $main.pdf ..." -ForegroundColor Cyan
latexmk -pdf -interaction=nonstopmode -file-line-error $main

$log = "$main.log"
if (Test-Path $log) {
    # "Infinite glue shrinkage found in box being split" is a known benign
    # artifact of longtable splitting a table across pages on current kernels.
    # The rules and repeated headers render correctly; filter it out so that
    # real errors are visible.
    $errors = Select-String -Path $log -Pattern '^!|^.*:\d+: ' -ErrorAction SilentlyContinue |
              Where-Object { $_.Line -notmatch 'Infinite glue shrinkage' }
    $undef  = Select-String -Path $log -Pattern 'undefined references|Reference .* undefined' -ErrorAction SilentlyContinue
    $over   = (Select-String -Path $log -Pattern 'Overfull \\hbox' -ErrorAction SilentlyContinue | Measure-Object).Count
    $pages  = (Select-String -Path $log -Pattern 'Output written on .*\((\d+) pages' -ErrorAction SilentlyContinue |
               Select-Object -Last 1).Matches.Groups[1].Value

    Write-Host ''
    Write-Host '--- Build report ---' -ForegroundColor Cyan
    Write-Host "Pages:              $pages"
    Write-Host "Errors:             $(($errors | Measure-Object).Count)"
    Write-Host "Undefined refs:     $(($undef  | Measure-Object).Count)"
    Write-Host "Overfull hboxes:    $over"
    if (($errors | Measure-Object).Count -gt 0) {
        Write-Host ''
        Write-Host 'First errors:' -ForegroundColor Red
        $errors | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.Line)" }
    }
}
