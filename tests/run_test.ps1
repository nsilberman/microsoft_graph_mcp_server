# PowerShell script to run login test

Write-Host "============================================================" -ForegroundColor Green
Write-Host "Running Login and Status Check Test" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

# Activate virtual environment if it exists
if (Test-Path ".venv\Scripts\Activate.ps1") {
    Write-Host "Activating virtual environment..." -ForegroundColor Yellow
    & .venv\Scripts\Activate.ps1
} else {
    Write-Host "Virtual environment not found, using system Python..." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Running test script..." -ForegroundColor Yellow
Write-Host ""

python test_login_complete.py

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Test completed!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green

Read-Host "Press Enter to exit"
