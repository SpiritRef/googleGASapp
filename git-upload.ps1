# 手動上傳 Git：自行決定何時提交與推送
$repo = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repo

Write-Host "=== Git Status ===" -ForegroundColor Cyan
git status

$msg = Read-Host "請輸入 commit 訊息（留空取消）"
if ([string]::IsNullOrWhiteSpace($msg)) {
    Write-Host "已取消，未提交。" -ForegroundColor Yellow
    exit
}

git add -A
git commit -m $msg

$push = Read-Host "是否立即 push？(y/n)"
if ($push -eq 'y' -or $push -eq 'Y') {
    git push
} else {
    Write-Host "已提交，稍後可手動 git push。" -ForegroundColor Green
}

Read-Host "按 Enter 結束"