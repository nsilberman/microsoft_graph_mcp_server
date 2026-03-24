@echo off
chcp 65001 >nul
echo ========================================
echo   Microsoft Graph MCP Server - Reinstall
echo ========================================
echo.

set "PROJECT_DIR=%~dp0"
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"

echo [1/2] Uninstalling old version...
uv tool uninstall microsoft-graph-mcp-server 2>nul

echo [2/2] Installing new version...
uv tool install --force "%PROJECT_DIR%"

echo.
echo ========================================
echo   Done! MCP Server has been reinstalled.
echo ========================================
echo.
echo Please restart Claude Desktop to apply changes.
pause
