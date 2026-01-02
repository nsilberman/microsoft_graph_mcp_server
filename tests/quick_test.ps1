# PowerShell script to run quick login test

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Quick Login Test" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Running python test_login_complete.py..." -ForegroundColor Yellow
Write-Host ""

python test_login_complete.py

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Test completed!" -ForegroundColor Cyan
Write-Host ""
Write-Host "Check the log file for details:" -ForegroundColor Yellow
Write-Host "  microsoft_graph_mcp_server/mcp_server_auth.log" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
