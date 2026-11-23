<#
.SYNOPSIS
  OfficeAI Windows 一键安装与集成脚本（按需启动模式，不做开机自启）

.DESCRIPTION
  1. 把 WPS 加载项安装为「本地离线」加载项，写入 %APPDATA%\kingsoft\wps\jsaddons
     - 加载项本身不依赖网关，因此可以承担「点击时自动拉起网关」的职责
     - 安装时把 OfficeAI 启动命令写入 main.js 的占位符
  2. 注册 officeai:// 协议（用于 Microsoft Office 用户一键启动网关）
  3. 注册 Microsoft Office Web Add-in 并信任本地 HTTPS 证书

  注意：不会写入任何开机自启项（Run / 启动文件夹 / 计划任务）。
#>
param(
  [switch]$SkipOffice,
  [switch]$SkipWps,
  [switch]$SkipProtocol,
  # 由谁来拉起网关：waker（Office COM 加载项，默认，不开机自启）
  #                   watcher（后台监听器，需要登录自启 + 常驻约 10MB）
  [ValidateSet('waker', 'watcher')]
  [string]$LaunchTrigger = 'waker',
  # 跳过自动启动组件（监听器 / 唤醒器）
  [switch]$SkipAutoStart
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== OfficeAI Windows 安装向导（按需启动） ===" -ForegroundColor Cyan

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

# 记录"从哪里启动 OfficeAI"，供 WPS 加载项按需拉起
$exeCandidates = @(
  (Join-Path $env:LOCALAPPDATA "Programs\OfficeAI\OfficeAI.exe"),
  (Join-Path $env:ProgramFiles "OfficeAI\OfficeAI.exe"),
  (Join-Path ${env:ProgramFiles(x86)} "OfficeAI\OfficeAI.exe"),
  (Join-Path $ScriptDir "OfficeAI.exe")
)
$OfficeAiExe = $exeCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if ($OfficeAiExe) {
  $env:OFFICEAI_LAUNCHER = $OfficeAiExe
  Write-Host "已检测到 OfficeAI.exe: $OfficeAiExe" -ForegroundColor Green
} else {
  $batPath = Join-Path $ScriptDir "Start-OfficeAI.bat"
  if (Test-Path $batPath) {
    $env:OFFICEAI_LAUNCHER = $batPath
    Write-Host "未检测到 OfficeAI.exe，按需启动将使用: $batPath" -ForegroundColor Yellow
  } else {
    Write-Host "未找到可用的启动程序，WPS 按钮只会提示手动启动。" -ForegroundColor Yellow
  }
}

$step = 0
$total = 0
if (-not $SkipWps) { $total++ }
if (-not $SkipProtocol) { $total++ }
if (-not $SkipAutoStart) { $total++ }
if (-not $SkipOffice) { $total++ }

if (-not $SkipWps) {
  $step++
  Write-Host "`n[$step/$total] 正在安装 WPS Office 本地离线加载项..." -ForegroundColor Yellow
  node $CliPath install-wps
}

if (-not $SkipProtocol) {
  $step++
  Write-Host "`n[$step/$total] 正在注册 officeai:// 一键启动协议..." -ForegroundColor Yellow
  node $CliPath install-protocol
}

if (-not $SkipAutoStart) {
  $step++
  if ($LaunchTrigger -eq 'watcher') {
    Write-Host "`n[$step/$total] 警告：安装后台监听器会注册登录自启并常驻运行。" -ForegroundColor Yellow
    node $CliPath install-watcher
  } else {
    Write-Host "`n[$step/$total] 正在安装 Office 唤醒器（COM 加载项，不开机自启、无常驻进程）..." -ForegroundColor Yellow
    node $CliPath install-waker
  }
}

if (-not $SkipOffice) {
  $step++
  Write-Host "`n[$step/$total] 正在注册 Microsoft Office 加载项并信任本地 HTTPS 证书..." -ForegroundColor Yellow
  node $CliPath install-office
}

Write-Host "`n安装完成！" -ForegroundColor Green
Write-Host "· WPS：重启 WPS 后点击顶部「OfficeAI 助手」，服务未启动时会自动拉起（不随系统开机自启）。"
if ($LaunchTrigger -eq 'watcher') {
  Write-Host "· Office：已安装后台监听器（登录自启 + 常驻）。如需取消：Uninstall-OfficeAI.bat，或改用 -LaunchTrigger waker"
} else {
  Write-Host "· Office：已安装唤醒器，打开 Word / Excel / PPT 时自动启动，登录时不运行任何进程。"
}
Write-Host "· 若自动启动组件未安装成功，可双击桌面「OfficeAI 启动」快捷方式手动拉起。"
Write-Host "· 关闭所有 Office / WPS 后，网关会在后台自动退出。"
