@echo off
title Microsoft Graph MCP Server - Login Test
color 0A

cls
echo ================================================================================
echo                      MCP SERVER LOGIN TEST
echo ================================================================================
echo.
echo This will test login and status check functionality
echo.
echo Press any key to start...
pause >nul

cls
echo ================================================================================
echo                      RUNNING LOGIN TEST
echo ================================================================================
echo.

python test_login_complete.py

echo.
echo ================================================================================
echo                           TEST COMPLETED
echo ================================================================================
echo.
echo Check the log file:
echo   microsoft_graph_mcp_server\mcp_server_auth.log
echo.
echo Also check these files in your user directory:
echo   %%USERPROFILE%%\.microsoft_graph_mcp_tokens.json
echo   %%USERPROFILE%%\.microsoft_graph_mcp_device_flows.json
echo.
pause
