# ============================================================
# push-app.ps1 - Day app (index.html, hoadon.html...) len GitHub
# Chay: chuot phai -> Run with PowerShell
#   hoac trong terminal: powershell -ExecutionPolicy Bypass -File "push-app.ps1"
# ============================================================

$ErrorActionPreference = "Continue"
$repo = "E:\Sieu du an\opc-nemthanh\du-an\DA05-nuoc-ion-kiem\web-kho-nuoc-ion"

Set-Location $repo
Write-Host "Thu muc: $repo" -ForegroundColor Cyan

# Nap tat ca thay doi
git add -A

# Commit (neu khong co gi moi thi bo qua, van push cac commit cu)
$msg = "update " + (Get-Date -Format "yyyy-MM-dd HH:mm")
git commit -m $msg 2>&1 | Out-Host

# Day len GitHub
Write-Host "`n>> Dang push len GitHub..." -ForegroundColor Yellow
git push 2>&1 | Out-Host

Write-Host "`n=== XONG ===" -ForegroundColor Green
Write-Host "1. Vao tab Actions tren GitHub, doi 'pages build and deployment' XANH."
Write-Host "   (Neu do 'try again later' -> bam Re-run all jobs.)"
Write-Host "2. Mo lai app them so moi o duoi URL, vi du:"
Write-Host "   https://dovantoanai-ui.github.io/kho-nuoc-ion/?v=99"
Write-Host "   https://dovantoanai-ui.github.io/kho-nuoc-ion/hoadon.html?v=99"

Read-Host "`nNhan Enter de dong"
