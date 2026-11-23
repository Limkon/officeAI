<#
.SYNOPSIS
  OfficeAI 唤醒器安装 / 注册 / 卸载（免管理员，仅当前用户）

.DESCRIPTION
  唤醒器是一个极小的 COM 加载项（OfficeAI.Waker.dll）。注册到
  HKCU\Software\Microsoft\Office\{Word,Excel,PowerPoint}\Addins 并设置
  LoadBehavior=3 后，Word / Excel / PowerPoint 启动时会自动加载它，
  由它检查网关是否运行、没有则拉起。

  它本身不提供任何 UI，也不常驻：Office 关闭后，网关的 LifecycleMonitor
  检测到所有 Office/WPS 进程消失即自动退出。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install-waker.ps1 -Action Install
  powershell -ExecutionPolicy Bypass -File install-waker.ps1 -Action Status
  powershell -ExecutionPolicy Bypass -File install-waker.ps1 -Action Uninstall
#>
param(
  [ValidateSet('Install', 'Uninstall', 'Status')]
  [string]$Action = 'Install',
  # 已编译好的 OfficeAI.Waker.dll 路径；默认在本脚本同级或 build 目录查找
  [string]$DllPath,
  # 自动编译（需要 .NET Framework 的 csc.exe，Windows 自带）
  [switch]$Build
)

$ErrorActionPreference = 'Stop'

$ProgId = 'OfficeAI.Waker'
$Clsid  = '{4F1D9C3A-8B27-4E6D-9C41-0A7B5E2D8F31}'
$Apps   = @('Word', 'Excel', 'PowerPoint')
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Resolve-Dll {
  param([string]$Explicit, [switch]$DoBuild)

  if ($Explicit -and (Test-Path $Explicit)) { return (Resolve-Path $Explicit).Path }

  foreach ($candidate in @(
    (Join-Path $ScriptDir 'OfficeAI.Waker.dll'),
    (Join-Path $ScriptDir 'build\OfficeAI.Waker.dll'),
    (Join-Path $ScriptDir '..\..\release\waker\OfficeAI.Waker.dll')
  )) {
    if (Test-Path $candidate) { return (Resolve-Path $candidate).Path }
  }

  if (-not $DoBuild) { return $null }

  # 用 Windows 自带的 .NET Framework 编译器，无需安装 .NET SDK
  $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
  if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
  }
  if (-not (Test-Path $csc)) {
    throw '未找到 csc.exe（.NET Framework 4.x）。请安装 .NET Framework 或使用 -DllPath 指定已编译的 DLL。'
  }

  $source = Join-Path $ScriptDir 'Waker.cs'
  if (-not (Test-Path $source)) { throw "未找到源码：$source" }

  $outDir = Join-Path $ScriptDir 'build'
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $out = Join-Path $outDir 'OfficeAI.Waker.dll'

  Write-Host "正在编译唤醒器..." -ForegroundColor Cyan
  & $csc /nologo /target:library /platform:anycpu /optimize+ `
    /out:"$out" `
    /reference:System.dll /reference:System.Core.dll "$source"

  if (-not (Test-Path $out)) { throw '编译失败。' }
  return $out
}

function Register-Dll {
  param([string]$Dll)

  # RegAsm 以 per-user 方式注册 COM 类（不需要管理员）
  $regasm = $null
  foreach ($candidate in @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe')
  )) {
    if (Test-Path $candidate) { $regasm = $candidate; break }
  }
  if (-not $regasm) { throw '未找到 RegAsm.exe（.NET Framework 4.x）。' }

  & $regasm /nologo /codebase /silent "$Dll"
  if ($LASTEXITCODE -ne 0) { throw "RegAsm 注册失败（exit=$LASTEXITCODE）。" }

  # per-user 注册项：HKCU 下同时补一份，避免任何 HKLM 依赖
  $progIdKey = "HKCU:\Software\Classes\$ProgId"
  New-Item -Path $progIdKey -Force | Out-Null
  Set-ItemProperty -Path $progIdKey -Name '(default)' -Value 'OfficeAI Office Waker'
  New-Item -Path "$progIdKey\CLSID" -Force | Out-Null
  Set-ItemProperty -Path "$progIdKey\CLSID" -Name '(default)' -Value $Clsid

  $clsidKey = "HKCU:\Software\Classes\CLSID\$Clsid"
  New-Item -Path $clsidKey -Force | Out-Null
  Set-ItemProperty -Path $clsidKey -Name '(default)' -Value 'OfficeAI Office Waker'
  New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name '(default)' -Value 'mscoree.dll'
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Class' -Value $ProgId
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Assembly' -Value ([Reflection.AssemblyName]::GetAssemblyName($Dll).FullName)
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319'
  Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'CodeBase' -Value ([Uri]$Dll).AbsoluteUri
  New-Item -Path "$clsidKey\InprocServer32\1.0.0.0" -Force | Out-Null
  Set-ItemProperty -Path "$clsidKey\InprocServer32\1.0.0.0" -Name 'Class' -Value $ProgId
  Set-ItemProperty -Path "$clsidKey\InprocServer32\1.0.0.0" -Name 'Assembly' -Value ([Reflection.AssemblyName]::GetAssemblyName($Dll).FullName)
  Set-ItemProperty -Path "$clsidKey\InprocServer32\1.0.0.0" -Name 'RuntimeVersion' -Value 'v4.0.30319'
  Set-ItemProperty -Path "$clsidKey\InprocServer32\1.0.0.0" -Name 'CodeBase' -Value ([Uri]$Dll).AbsoluteUri
  New-Item -Path "$clsidKey\ProgId" -Force | Out-Null
  Set-ItemProperty -Path "$clsidKey\ProgId" -Name '(default)' -Value $ProgId

  # 让 Office 在启动时加载（LoadBehavior = 3）
  foreach ($app in $Apps) {
    foreach ($root in @("HKCU:\Software\Microsoft\Office\$app\Addins", "HKCU:\Software\Microsoft\Office\16.0\$app\Addins")) {
      $key = Join-Path $root $ProgId
      New-Item -Path $key -Force | Out-Null
      Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'OfficeAI 唤醒器（按需启动本地服务）'
      Set-ItemProperty -Path $key -Name 'Description' -Value 'OfficeAI 在 Word/Excel/PPT 启动时自动拉起本地网关，关闭后自动退出。'
      Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord
      Set-ItemProperty -Path $key -Name 'CommandLineSafe' -Value 1 -Type DWord
    }
  }
}

function Unregister-Dll {
  $regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
  if (-not (Test-Path $regasm)) {
    $regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
  }
  if (Test-Path $regasm) {
    if ($DllPath -and (Test-Path $DllPath)) {
      & $regasm /nologo /unregister /silent "$DllPath" | Out-Null
    }
  }

  foreach ($app in $Apps) {
    foreach ($root in @("HKCU:\Software\Microsoft\Office\$app\Addins", "HKCU:\Software\Microsoft\Office\16.0\$app\Addins")) {
      $key = Join-Path $root $ProgId
      if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
    }
  }
  foreach ($key in @("HKCU:\Software\Classes\CLSID\$Clsid", "HKCU:\Software\Classes\$ProgId")) {
    if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
  }
}

function Get-Status {
  $result = [ordered]@{ progId = $ProgId; clsid = $Clsid; apps = @() }
  foreach ($app in $Apps) {
    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\$ProgId"
    $loadBehavior = $null
    if (Test-Path $key) {
      $loadBehavior = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).LoadBehavior
    }
    $result.apps += [pscustomobject]@{ app = $app; registered = [bool](Test-Path $key); loadBehavior = $loadBehavior }
  }
  $result | ConvertTo-Json -Depth 4
}

switch ($Action) {
  'Install' {
    $dll = Resolve-Dll -Explicit $DllPath -DoBuild:$Build
    if (-not $dll) { throw '未找到 OfficeAI.Waker.dll，请使用 -Build 现场编译或 -DllPath 指定路径。' }
    Register-Dll -Dll $dll
    Write-Host "唤醒器已安装：$dll" -ForegroundColor Green
    Write-Host "Word / Excel / PowerPoint 下次启动时会自动检查并拉起 OfficeAI 本地服务。"
    Write-Host "可通过「文件 → 选项 → 加载项 → COM 加载项」查看或临时禁用。"
    exit 0
  }
  'Uninstall' {
    Unregister-Dll
    Write-Host '唤醒器已卸载（Office 不再自动启动 OfficeAI）。' -ForegroundColor Green
    exit 0
  }
  'Status' {
    Get-Status
    exit 0
  }
}
