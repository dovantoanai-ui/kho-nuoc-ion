# Day app kho DA05 len GitHub -> lay link web (GitHub Pages)
# Chay: chuot phai file nay -> Run with PowerShell  (hoac dan noi dung vao cua so PowerShell)

$path = "E:\Sieu du an\opc-nemthanh\du-an\DA05-nuoc-ion-kiem\web-kho-nuoc-ion"
$user = "dovantoanai-ui"
$repo = "kho-nuoc-ion"

Set-Location $path
if (-not (Test-Path ".git")) { git init | Out-Null }
git add .
git commit -m "App kho nuoc ion kiem DA05" 2>$null
git branch -M main

if (Get-Command gh -ErrorAction SilentlyContinue) {
    # Co GitHub CLI -> tao repo + push + bat Pages tu dong
    gh repo create "$user/$repo" --public --source . --remote origin --push
    gh api -X POST "repos/$user/$repo/pages" -f "source[branch]=main" -f "source[path]=/" 2>$null
    Write-Host ""
    Write-Host "XONG. Link web app: https://$user.github.io/$repo/" -ForegroundColor Green
} else {
    # Khong co gh CLI -> can tao repo rong truoc tren web
    if (-not (git remote | Select-String "origin")) {
        git remote add origin "https://github.com/$user/$repo.git"
    }
    git push -u origin main
    Write-Host ""
    Write-Host "Da push code." -ForegroundColor Green
    Write-Host "Neu bao 'repository not found': vao https://github.com/new tao repo ten '$repo' (Public, KHONG add README) roi chay lai file nay." -ForegroundColor Yellow
    Write-Host "Sau do bat Pages: repo -> Settings -> Pages -> Deploy from branch -> main / root -> Save." -ForegroundColor Yellow
    Write-Host "Link web app: https://$user.github.io/$repo/" -ForegroundColor Green
}
