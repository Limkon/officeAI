<#
.SYNOPSIS
  OfficeAI 后台监听器安装 / 卸载 / 状态查询（免管理员）

.DESCRIPTION
  监听器（OfficeAI.Watcher.exe）是一个约 10MB 内存的极轻量常驻进程：
  每 2 秒检查一次是否有 Word / Excel / PowerPoint / WPS 在运行，发现办公
  软件启动且网关未运行时，把网关拉起。

  与 Office 唤醒器（COM 加载项）的区别：
    Watcher → 常驻进程，Office 启动路径上零开销，绝不会被"慢加载项"禁用
    Waker   → 无常驻进程，由 Office 在宿主进程内加载，占用为 0

  安装后会在登录时自启（计划任务，无需管理员）。可用 -Action Uninstall 移除。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install-watcher.ps1 -Action Install -Build
  powershell -ExecutionPolicy Bypass -File install-watcher.ps1 -Action Status
  powershell -ExecutionPolicy Bypass -File install-watcher.ps1 -Action Uninstall
#>
param(
  [ValidateSet('Install', 'Uninstall', 'Status')]
  [string]$Action = 'Install',
  [string]$ExePath,
  [switch]$Build
)

$ErrorActionPreference = 'Stop'
$TaskName = 'OfficeAI Watcher'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExeName = 'OfficeAI.Watcher.exe'

function Build-Exe {
  # 优先 Roslyn 编译器（支持新版 C#），否则退回 .NET Framework 自带的 C# 5 编译器。
  # 源码本身已按 C# 5 编写（无 out var / 字符串插值 / ?.），两条路都能编。
  $csc = $null
  foreach ($candidate in @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
  )) {
    if (Test-Path $candidate) { $csc = $candidate; break }
  }
  if (-not $csc) {
    throw '未找到 csc.exe（.NET Framework 4.x）。'
  }
  $source = Join-Path $ScriptDir 'Watcher.cs'
  if (-not (Test-Path $source)) { throw "未找到源码：$source" }

  $outDir = Join-Path $ScriptDir 'build'
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $exe = Join-Path $outDir $ExeName

  Write-Host '正在编译后台监听器...' -ForegroundColor Cyan
  # /target:winexe 表示没有控制台窗口
  & $csc /nologo /target:winexe /platform:anycpu /optimize+ /out:"$exe" `
    /reference:System.dll /reference:System.Core.dll "$source"
  if (-not (Test-Path $exe)) { throw '编译失败。' }
  return $exe
}

function Resolve-Exe {
  param([string]$Explicit, [switch]$DoBuild)
  if ($Explicit -and (Test-Path $Explicit)) { return (Resolve-Path $Explicit).Path }
  foreach ($candidate in @(
    (Join-Path $ScriptDir $ExeName),
    (Join-Path $ScriptDir "build\$ExeName")
  )) {
    if (Test-Path $candidate) { return (Resolve-Path $candidate).Path }
  }
  if ($DoBuild) { return (Build-Exe) }
  return $null
}

function Install-StartupTask {
  param([string]$Exe)

  # 存一份到用户目录，避免源码目录被移动后失效
  $installDir = Join-Path $env:LOCALAPPDATA 'OfficeAI\Watcher'
  New-Item -ItemType Directory -Force -Path $installDir | Out-Null
  $target = Join-Path $installDir $ExeName
  Copy-Item -Path $Exe -Destination $target -Force

  $action = New-ScheduledTaskAction -Execute $target -Argument '--watch'
  $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
  $settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Description 'OfficeAI：检测到 Word/Excel/PPT/WPS 启动时自动拉起本地服务' -Force | Out-Null

  Start-ScheduledTask -TaskName $TaskName
  return $target
}

function Remove-StartupTask {
  $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  if ($task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  }
  # 结束可能残留的进程（按可执行文件路径精确匹配，不误杀其他 dotnet 进程）
  Get-Process -Name 'OfficeAI.Watcher' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  $installDir = Join-Path $env:LOCALAPPDATA 'OfficeAI\Watcher'
  if (Test-Path $installDir) { Remove-Item -Path $installDir -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-WatcherStatus {
  $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
  $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
  $proc = Get-Process -Name 'OfficeAI.Watcher' -ErrorAction SilentlyContinue
  $exeCandidates = @(
    (Join-Path $ScriptDir $ExeName),
    (Join-Path $ScriptDir "build\$ExeName")
  ) | Where-Object { Test-Path $_ }

  [ordered]@{
    taskName = $TaskName
    exePath = if ($exeCandidates) { (Resolve-Path $exeCandidates[0]).Path } else { $null }
    registered = [bool]$task
    state = if ($task) { [string]$task.State } else { 'Absent' }
    running = [bool]$proc
    pid = if ($proc) { $proc.Id } else { $null }
    memoryMB = if ($proc) { [math]::Round($proc.WorkingSet64 / 1MB, 1) } else { $null }
    lastRunTime = if ($info) { [string]$info.LastRunTime } else { $null }
    startupBudgetMs = 0
  } | ConvertTo-Json -Depth 4
}

# -Build 对任何 Action 都生效（CI 会以 -Action Status -Build 只编译不注册）
if ($Build) {
  $builtExe = Build-Exe
  Write-Host "已编译：$builtExe" -ForegroundColor Green
}

switch ($Action) {
  'Install' {
    $exe = Resolve-Exe -Explicit $ExePath -DoBuild:$Build
    if (-not $exe) { throw "未找到 $ExeName，请使用 -Build 编译或 -DllPath/-ExePath 指定路径。" }
    $target = Install-StartupTask -Exe $exe
    Write-Host "后台监听器已安装：$target" -ForegroundColor Green
    Write-Host '已注册登录自启计划任务（当前用户，无需管理员）。'
    Write-Host '请保持「启动方式」为「按需启动」且触发方式为「后台监听器」。'
    exit 0
  }
  'Uninstall' {
    Remove-StartupTask
    Write-Host '后台监听器已卸载（登录自启已移除，进程已结束）。' -ForegroundColor Green
    exit 0
  }
  'Status' {
    Get-WatcherStatus
    exit 0
  }
}
