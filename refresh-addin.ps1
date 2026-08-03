# DeepSeek 文档助手 - 刷新加载项缓存脚本
param(
    [switch]$Auto
)
$ErrorActionPreference = "Stop"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  DeepSeek 文档助手 - 刷新加载项缓存" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$word = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($word) {
    Write-Host ""
    Write-Host "检测到 Word 正在运行。请先保存所有文档。" -ForegroundColor Yellow
    if ($Auto) {
        Write-Host "自动模式：关闭 Word 并继续。"
        $ans = "Y"
    } else {
        $ans = Read-Host "是否自动关闭 Word 并继续？(Y/N)"
    }
    if ($ans -notmatch "^[Yy]") {
        Write-Host "已取消。请保存文档后重新运行本脚本。" -ForegroundColor Yellow
        if (-not $Auto) { Read-Host "按回车退出" }
        exit 1
    }
    Write-Host "正在关闭 Word..."
    Stop-Process -Name WINWORD -Force
    Start-Sleep -Seconds 2
}

$cache = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
if (Test-Path $cache) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backup = $cache + "_backup_" + $stamp
    Rename-Item -LiteralPath $cache -NewName (Split-Path $backup -Leaf)
    Write-Host ("缓存已备份到: " + $backup) -ForegroundColor Green
    Write-Host "缓存已清除，Word 下次打开会自动重新下载最新版加载项。" -ForegroundColor Green
} else {
    Write-Host "未发现 WEF 缓存，无需清理。" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "正在重新打开 Word..."
$winword = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if (Test-Path $winword) {
    Start-Process -FilePath $winword
} else {
    Start-Process winword
}

Write-Host ""
Write-Host "完成！打开文档后：插入 -> 加载项 -> 我的加载项 -> 共享文件夹 -> DeepSeek 文档助手" -ForegroundColor Green
Write-Host "发送一条消息后，查看底部状态栏是否显示「已附带文档上下文」。" -ForegroundColor Green
if (-not $Auto) { Read-Host "按回车退出" }
