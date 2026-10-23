# 小恐龙公文排版助手 —— 本地卸载脚本
#
# 作用：清理 install.ps1 写入的 Word 加载项与 WPS 白名单条目。
#
# 用法：双击同目录下的「卸载.bat」，或
#       powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1

$ErrorActionPreference = 'Stop'

$keyName = '小恐龙工作室.小恐龙公文排版助手'

$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\$keyName"
if (Test-Path $addinKey) {
    Remove-Item $addinKey -Recurse -Force
    Write-Host "已删除 Word 加载项：$keyName"
}

$wl = 'HKCU:\Software\Kingsoft\Office\WPS\AddinsWL'
if (Test-Path $wl) {
    $existing = Get-ItemProperty -Path $wl -Name $keyName -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-ItemProperty -Path $wl -Name $keyName
        Write-Host "已从 WPS 白名单移除：$keyName"
    }
}

$legacy = 'HKCU:\Software\Microsoft\Office\Word\Addins\公文排版'
if (Test-Path $legacy) {
    Remove-Item $legacy -Recurse -Force
    Write-Host '已删除 VSTOInstaller 注册项：公文排版'
}

Write-Host ''
Write-Host '卸载完成。请重启 Word / WPS。'
