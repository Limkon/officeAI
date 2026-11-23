@echo off
chcp 65001 >nul
title OfficeAI 本地服务（按需启动）
setlocal
set ROOT_DIR=%~dp0
if exist "%ROOT_DIR%gateway\cli.js" (
  set CLI_PATH=%ROOT_DIR%gateway\cli.js
  set OFFICEAI_ASSETS_DIR=%ROOT_DIR%assets
) else (
  set CLI_PATH=%ROOT_DIR%..\..\packages\gateway\dist\cli.js
)

rem --launched-by 只用于日志/诊断：记录是谁按需拉起了网关
if "%~1"=="" (
  node "%CLI_PATH%" --launched-by manual
) else (
  node "%CLI_PATH%" --launched-by manual --launch "%~1"
)
pause
