<#
.SYNOPSIS
  OfficeAI Windows 卸载脚本（从 Microsoft Office 与 WPS Office 移除加载项）
#>
$ErrorActionPreference = "Continue"

$WpsAddons = Join-Path $env:APPDATA "kingsoft\wps\jsaddons"
$LocalAddon = Join-Path $WpsAddons "officeai_1.0.0"
if (Test-Path $LocalAddon) {
  Remove-Item -Path $LocalAddon -Recurse -Force
  Write-Host "已删除本地离线加载项: $LocalAddon"
}
foreach ($xmlName in @("publish.xml", "jsplugins.xml")) {
  $xmlFile = Join-Path $WpsAddons $xmlName
  if (Test-Path $xmlFile) {
    $content = Get-Content $xmlFile -Raw -Encoding UTF8
    $updated = [regex]::Replace($content, '\s*<jsplugin\b[^>]*name="officeai"[^>]*/>', '', 'IgnoreCase')
    Set-Content -Path $xmlFile -Value $updated -Encoding UTF8
    Write-Host "已清理 WPS 配置: $xmlFile"
  }
}

$RegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
foreach ($AddinId in @("b8d4c3e2-71a0-4f88-9c3d-0f1ce2026a00", "b8d4c3e2-71a0-4f88-9c3d-0f1ce2026a001")) {
  if (Test-Path $RegPath) {
    Remove-ItemProperty -Path $RegPath -Name $AddinId -ErrorAction SilentlyContinue
  }
}
Write-Host "已清理 Microsoft Office 注册项。"

# 卸载 Office 唤醒器（COM 加载项）
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$WakerScript = @(
  (Join-Path $ScriptDir "waker\install-waker.ps1"),
  (Join-Path $ScriptDir "..\..\windows\waker\install-waker.ps1"),
  (Join-Path $env:LOCALAPPDATA "Programs\OfficeAI\resources\assets\waker\install-waker.ps1")
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if ($WakerScript) {
  & powershell -NoProfile -ExecutionPolicy Bypass -File $WakerScript -Action Uninstall
} else {
  # 没有脚本时直接清注册表（与 install-waker.ps1 的卸载逻辑等价）
  foreach ($app in @("Word", "Excel", "PowerPoint")) {
    foreach ($root in @("HKCU:\Software\Microsoft\Office\$app\Addins", "HKCU:\Software\Microsoft\Office\16.0\$app\Addins")) {
      $key = Join-Path $root "OfficeAI.Waker"
      if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
    }
  }
  foreach ($key in @(
    "HKCU:\Software\Classes\CLSID\{4F1D9C3A-8B27-4E6D-9C41-0A7B5E2D8F31}",
    "HKCU:\Software\Classes\OfficeAI.Waker"
  )) {
    if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
  }
  Write-Host "已清理 Office 唤醒器注册项。"
}

# 清理按需启动使用的 officeai:// 协议与桌面快捷方式
$ProtocolKey = "HKCU:\Software\Classes\officeai"
if (Test-Path $ProtocolKey) {
  Remove-Item -Path $ProtocolKey -Recurse -Force
  Write-Host "已移除 officeai:// 协议注册。"
}
foreach ($Shortcut in @(
  (Join-Path ([Environment]::GetFolderPath('Desktop')) "OfficeAI 启动.url"),
  (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\OfficeAI 启动.url")
)) {
  if (Test-Path $Shortcut) {
    Remove-Item -Path $Shortcut -Force
    Write-Host "已删除快捷方式: $Shortcut"
  }
}

Write-Host "卸载完成。" -ForegroundColor Green
