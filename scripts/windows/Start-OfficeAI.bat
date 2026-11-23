@echo off
chcp 65001 >nul
title OfficeAI 启动
setlocal
rem 手动 / 排障入口：直接启动已安装的 OfficeAI（不注册任何自启项）。
rem 平时无需运行本脚本——打开 Word / Excel / PPT / WPS 时会自动启动。

set "EXE=%LOCALAPPDATA%\Programs\OfficeAI\OfficeAI.exe"
if not exist "%EXE%" set "EXE=%ProgramFiles%\OfficeAI\OfficeAI.exe"
if not exist "%EXE%" set "EXE=%ProgramFiles(x86)%\OfficeAI\OfficeAI.exe"
if not exist "%EXE%" set "EXE=%~dp0OfficeAI.exe"

if not exist "%EXE%" (
  echo 未找到 OfficeAI.exe，请先运行 OfficeAI Setup 安装包。
  pause
  exit /b 1
)

if "%~1"=="" (
  start "" "%EXE%" --launched-by manual
) else (
  start "" "%EXE%" --launched-by manual --launch "%~1"
)
