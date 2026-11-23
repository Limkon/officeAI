@echo off
chcp 65001 >nul
rem 用法: Install-OfficeAI.bat [-LaunchTrigger watcher|waker] [-SkipOffice] [-SkipWps]
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-OfficeAI.ps1" %*
pause
