# DeepSeek 文档助手 - Word 桌面版一键安装脚本
$ErrorActionPreference = 'Stop'

function Write-Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-WarnMsg($msg) { Write-Host "    [!!] $msg" -ForegroundColor Yellow }

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  DeepSeek 文档助手 - Word 桌面版一键安装" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# 1. 管理员权限检查
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-WarnMsg "需要管理员权限，请右键 install.bat 选择「以管理员身份运行」"
    Read-Host "按回车退出"
    exit 1
}

# 2. 检测 Word
$winword = Get-ChildItem "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE","C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE","C:\Program Files\Microsoft Office\Office16\WINWORD.EXE" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($winword) {
    Write-Ok ("检测到 Word: " + $winword.FullName)
} else {
    Write-WarnMsg "未在常见路径检测到 Word，安装仍会继续（请确认已安装 Microsoft 365/Office）"
}

# 3. 获取 manifest（优先本地，其次从 GitHub 下载）
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$localManifest = Join-Path $scriptDir "manifest.xml"
$destDir = "C:\WordAddins"
$destManifest = Join-Path $destDir "manifest.xml"
$manifestContent = $null

if (Test-Path $localManifest) {
    Write-Step "使用脚本同目录下的 manifest.xml"
    $manifestContent = Get-Content -Raw -Encoding UTF8 $localManifest
} else {
    Write-Step "本地没有 manifest.xml，尝试从 GitHub 下载"
    $urls = @(
        "https://tap-apia.github.io/word-deepseek-assistant/manifest.xml",
        "https://raw.githubusercontent.com/TAP-APIA/word-deepseek-assistant/main/manifest.xml"
    )
    foreach ($u in $urls) {
        try {
            $wc = New-Object System.Net.WebClient
            $manifestContent = $wc.DownloadString($u)
            Write-Ok ("下载成功: " + $u)
            break
        } catch {
            Write-WarnMsg ("下载失败: " + $u)
        }
    }
}

if (-not $manifestContent -or $manifestContent -notmatch "<OfficeApp") {
    Write-WarnMsg "无法获得有效的 manifest.xml：请把 manifest.xml 放到脚本同一目录后重试"
    Read-Host "按回车退出"
    exit 1
}

# 4. 创建目录并写入 manifest
Write-Step "准备目录 $destDir"
New-Item -ItemType Directory -Path $destDir -Force | Out-Null
[System.IO.File]::WriteAllText($destManifest, $manifestContent, (New-Object System.Text.UTF8Encoding($false)))
Write-Ok "manifest.xml 已就位: $destManifest"

# 5. 创建网络共享（Word 要求共享文件夹的 UNC 路径）
Write-Step "创建网络共享 WordAddins"
$shareName = "WordAddins"
$existingShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if ($existingShare) {
    if ($existingShare.Path -eq $destDir) {
        Write-Ok "共享已存在且指向正确路径"
    } else {
        Write-WarnMsg ("检测到同名共享指向其他路径: " + $existingShare.Path + "，将修正")
        Remove-SmbShare -Name $shareName -Force
        New-SmbShare -Name $shareName -Path $destDir -ReadAccess Everyone | Out-Null
        Write-Ok "共享已重建"
    }
} else {
    New-SmbShare -Name $shareName -Path $destDir -ReadAccess Everyone | Out-Null
    Write-Ok ("共享已创建: \\127.0.0.1\" + $shareName)
}

# 6. 注册受信任的加载项目录（注册表）
Write-Step "注册受信任的加载项目录"
$unc = "\\127.0.0.1\$shareName"
$officeVersions = @("16.0", "15.0")
$registered = $false
foreach ($ver in $officeVersions) {
    $base = "HKCU:\Software\Microsoft\Office\$ver\WEF\TrustedCatalogs"
    if (-not (Test-Path $base)) { continue }
    $existingEntry = Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object { Get-ItemProperty $_.PSPath } | Where-Object { $_.Url -eq $unc }
    if ($existingEntry) {
        Write-Ok ("Word $ver 已存在登记: " + $unc)
        $registered = $true
    } else {
        $guid = "{" + [guid]::NewGuid().ToString() + "}"
        $keyPath = Join-Path $base $guid
        New-Item -Path $keyPath -Force | Out-Null
        New-ItemProperty -Path $keyPath -Name "Id" -Value $guid -PropertyType String | Out-Null
        New-ItemProperty -Path $keyPath -Name "Url" -Value $unc -PropertyType String | Out-Null
        New-ItemProperty -Path $keyPath -Name "Flags" -Value 1 -PropertyType DWord | Out-Null
        Write-Ok ("Word $ver 已登记: " + $unc)
        $registered = $true
    }
}
if (-not $registered) {
    Write-WarnMsg "未找到 Office 的 WEF 注册表路径，请确认已安装 Office"
}

# 7. 最终验证
Write-Step "最终验证"
$testPath = "\\127.0.0.1\$shareName\manifest.xml"
if (Test-Path $testPath) {
    Write-Ok ("共享可访问: " + $testPath)
} else {
    Write-WarnMsg "共享访问验证失败，请检查防火墙或共享设置"
}

# 8. 完成提示
$wordRunning = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "安装完成！" -ForegroundColor Green
if ($wordRunning) {
    Write-Host "检测到 Word 正在运行：请先保存文档，完全退出 Word 后重新打开。" -ForegroundColor Yellow
} else {
    Write-Host "请打开 Word 完成最后两步：" -ForegroundColor White
}
Write-Host "1. 插入 -> 加载项 -> 我的加载项" -ForegroundColor White
Write-Host "2. 顶部选择「共享文件夹」-> DeepSeek 文档助手 -> 添加" -ForegroundColor White
Write-Host "首次打开后在侧边栏点 ⚙ 填写 DeepSeek API Key。" -ForegroundColor White
Read-Host "按回车退出"
