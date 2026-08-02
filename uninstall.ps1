# DeepSeek 文档助手 - 卸载脚本
$ErrorActionPreference = 'Stop'

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  DeepSeek 文档助手 - 卸载" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "需要管理员权限，请右键 uninstall.bat 选择「以管理员身份运行」" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

$unc = "\\127.0.0.1\WordAddins"

# 1. 移除受信任目录注册表登记
Write-Host ""
Write-Host "==> 移除受信任目录注册表登记" -ForegroundColor Cyan
$removed = $false
foreach ($ver in @("16.0", "15.0")) {
    $base = "HKCU:\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs"
    if (Test-Path $base) {
        Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath
            if ($p.Url -eq $unc) {
                Remove-Item -LiteralPath $_.PSPath -Recurse -Force
                Write-Host ("    已移除 (Word $ver): " + $p.Url) -ForegroundColor Green
                $removed = $true
            }
        }
    }
}
if (-not $removed) { Write-Host "    未找到需要移除的登记" -ForegroundColor Yellow }

# 2. 移除网络共享
Write-Host ""
Write-Host "==> 移除网络共享 WordAddins" -ForegroundColor Cyan
$share = Get-SmbShare -Name WordAddins -ErrorAction SilentlyContinue
if ($share) {
    Remove-SmbShare -Name WordAddins -Force
    Write-Host "    共享已移除" -ForegroundColor Green
} else {
    Write-Host "    共享不存在" -ForegroundColor Yellow
}

# 3. 询问是否删除文件夹
if (Test-Path C:\WordAddins) {
    Write-Host ""
    $ans = Read-Host "是否同时删除 C:\WordAddins 文件夹？(Y/N)"
    if ($ans -match "^[Yy]") {
        Remove-Item -LiteralPath C:\WordAddins -Recurse -Force
        Write-Host "    C:\WordAddins 已删除" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "卸载完成。请重启 Word（若已打开）以生效。" -ForegroundColor Green
Read-Host "按回车退出"
