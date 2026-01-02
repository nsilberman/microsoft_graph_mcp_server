@echo off
REM Simple script to run MCP server with logging to file

echo Starting MCP Server with logging...
echo Logs will be saved to: mcp_server.log
echo.

python -m microsoft_graph_mcp_server.main 2>&1

pause
