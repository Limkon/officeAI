@echo off
chcp 65001 >nul
title OfficeAI 智能办公网关
setlocal
set ROOT_DIR=%~dp0
if exist "%ROOT_DIR%gateway\cli.js" (
  set CLI_PATH=%ROOT_DIR%gateway\cli.js
  set OFFICEAI_ASSETS_DIR=%ROOT_DIR%assets
) else (
  set CLI_PATH=%ROOT_DIR%..\..\packages\gateway\dist\cli.js
)

echo 正在启动 OfficeAI 本地网关...
echo 控制中心: http://127.0.0.1:18080/dashboard
echo 按 Ctrl+C 停止服务
echo.
start "" "http://127.0.0.1:18080/dashboard"
node "%CLI_PATH%"
pause
