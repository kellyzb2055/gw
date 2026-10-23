# 小恐龙公文排版助手 —— 本地安装脚本
#
# 作用：按官方安装包的注册方式写入 Word 加载项与 WPS 白名单。
#       VSTOInstaller（双击 .vsto）只会写 Word 的「公文排版」项，
#       不会写 WPS 白名单，因此在 WPS 里不显示。
#
# 用法：双击同目录下的「安装.bat」，或
#       powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1

$ErrorActionPreference = 'Stop'

$dir = $PSScriptRoot
$vsto = Join-Path $dir '公文排版.vsto'
if (-not (Test-Path $vsto)) {
    throw "未在 $dir 找到 公文排版.vsto，请把本脚本放在解压目录根下运行。"
}

$keyName = '小恐龙工作室.小恐龙公文排版助手'
$manifestUrl = 'file:///' + ($vsto -replace '\\', '/') + '|vstolocal'

# 1) Word 加载项注册项：键名必须与 WPS 白名单中的条目一致
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$keyName"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'Description'  -Value '小恐龙公文助手' -Type String
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value '公文排版助手'   -Type String
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3               -Type DWord
Set-ItemProperty -Path $addinKey -Name 'Manifest'     -Value $manifestUrl    -Type String
Write-Host "已写入 Word 加载项：$keyName"
Write-Host "  Manifest = $manifestUrl"

# 2) WPS 白名单：WPS 只加载 AddinsWL 中登记过的项，缺这一条在 WPS 里不显示
$wl = 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL'
if (-not (Test-Path $wl)) { New-Item -Path $wl -Force | Out-Null }
Set-ItemProperty -Path $wl -Name $keyName -Value '' -Type String
Write-Host "已写入 WPS 白名单：$keyName"

# 3) 停用 VSTOInstaller 自动创建的「公文排版」项，避免同一插件被加载两次
$legacy = 'HKCU:\Software\Microsoft\Office\Word\Addins\公文排版'
if (Test-Path $legacy) {
    Set-ItemProperty -Path $legacy -Name 'LoadBehavior' -Value 0 -Type DWord
    Write-Host '已停用 VSTOInstaller 注册的重复项：公文排版'
}

Write-Host ''
Write-Host '安装完成。请完全退出 Word 和 WPS（含托盘进程）后重新打开。'
