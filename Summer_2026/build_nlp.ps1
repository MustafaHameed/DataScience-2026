# Build the Summer 2026 NLP handout.
#   .\build_nlp.ps1          normal build
#   .\build_nlp.ps1 -Clean   remove aux files first
param([switch]$Clean)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$main = 'NLP_Handout'

if ($Clean) {
    Remove-Item -ErrorAction SilentlyContinue "$main.aux","$main.log","$main.out","$main.toc"
    Write-Host 'Cleaned auxiliary files.' -ForegroundColor Yellow
}

Write-Host "Building $main.pdf ..." -ForegroundColor Cyan
pdflatex -interaction=nonstopmode -file-line-error $main | Out-Null
pdflatex -interaction=nonstopmode -file-line-error $main | Out-Null

$log = "$main.log"
if (Test-Path $log) {
    $errors = Select-String -Path $log -Pattern '^!|^.*:\d+: ' -ErrorAction SilentlyContinue
    $pages  = (Select-String -Path $log -Pattern 'Output written on .*\((\d+) pages' -ErrorAction SilentlyContinue |
               Select-Object -Last 1).Matches.Groups[1].Value

    Write-Host ''
    Write-Host '--- Build report ---' -ForegroundColor Cyan
    Write-Host "Pages:  $pages"
    Write-Host "Errors: $(($errors | Measure-Object).Count)"
    if (($errors | Measure-Object).Count -gt 0) {
        Write-Host ''
        Write-Host 'First errors:' -ForegroundColor Red
        $errors | Select-Object -First 10 | ForEach-Object { Write-Host "  $($_.Line)" }
    }
}
