<#
.SYNOPSIS
  OfficeAI Windows 一键安装与集成脚本（支持 Microsoft Office 与 WPS Office）
#>
param(
  [switch]$SkipOffice,
  [switch]$SkipWps
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== OfficeAI Windows 安装向导 ===" -ForegroundColor Cyan

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Write-Host "未检测到 Node.js，请先安装 Node.js >= 20 (https://nodejs.org) 或使用 OfficeAI-Setup.exe 安装包。" -ForegroundColor Red
  exit 1
}

$Candidates = @(
  (Join-Path $ScriptDir "gateway\cli.js"),
  (Join-Path $ScriptDir "..\gateway\cli.js"),
  (Join-Path $ScriptDir "..\..\packages\gateway\dist\cli.js")
)

$CliPath = $Candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $CliPath) {
  Write-Host "未找到已编译的 Gateway，请先执行 npm run build。" -ForegroundColor Red
  exit 1
}

$PortableAssets = Join-Path $ScriptDir "assets"
if (Test-Path $PortableAssets) {
  $env:OFFICEAI_ASSETS_DIR = $PortableAssets
}

if (-not $SkipWps) {
  Write-Host "`n[1/2] 正在注册 WPS Office JS 加载项..." -ForegroundColor Yellow
  node $CliPath install-wps
}

if (-not $SkipOffice) {
  Write-Host "`n[2/2] 正在注册 Microsoft Office 加载项并信任本地 HTTPS 证书..." -ForegroundColor Yellow
  node $CliPath install-office
}

Write-Host "`n安装完成！双击 Start-OfficeAI.bat 启动服务，然后打开 Word/Excel/PPT 或 WPS 即可使用。" -ForegroundColor Green
