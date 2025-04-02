# Uninstall.ps1
# Uninstallation script for the GCI-Notion365 Reporting Tool

# Configuration
$DataDir = "$env:ProgramData\GCI-Notion365"
$LogFile = Join-Path -Path $env:TEMP -ChildPath "GCI-Notion365-Uninstall.log"

# Start logging
Start-Transcript -Path $LogFile -Force

Write-Host "GCI-Notion365 Uninstallation" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan

# Remove data directories
if (Test-Path -Path $DataDir) {
    Write-Host "Removing data directory: $DataDir..." -ForegroundColor Yellow
    
    $keepData = Read-Host "Do you want to keep the report data? (y/n)"
    if ($keepData.ToLower() -eq "y") {
        Write-Host "Keeping data directory as requested." -ForegroundColor Green
    } else {
        try {
            Remove-Item -Path $DataDir -Recurse -Force
            Write-Host "Data directory removed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to remove data directory: $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "Data directory not found: $DataDir" -ForegroundColor Yellow
}

# Clean up any temporary files
$tempPattern = "$env:TEMP\GCI-Notion365*.temp"
if (Test-Path -Path $tempPattern) {
    Write-Host "Removing temporary files..." -ForegroundColor Yellow
    try {
        Remove-Item -Path $tempPattern -Force
        Write-Host "Temporary files removed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to remove temporary files: $_" -ForegroundColor Red
    }
}

# Completion message
Write-Host ""
Write-Host "GCI-Notion365 uninstallation completed!" -ForegroundColor Green
Write-Host "To completely remove GCI-Notion365, you can now delete the script directory manually." -ForegroundColor Yellow
Write-Host "Uninstallation log saved to: $LogFile" -ForegroundColor Cyan

Stop-Transcript