# Logging.ps1
# Logging functions for GCI-Notion365 Reporting Tool

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Make sure the log file path is available
    if (-not (Test-Path -Path (Split-Path -Path $script:logFile -Parent))) {
        New-Item -Path (Split-Path -Path $script:logFile -Parent) -ItemType Directory -Force | Out-Null
    }
    
    Add-Content -Path $script:logFile -Value $logEntry
    
    if ($Level -eq "ERROR") {
        Write-Host $logEntry -ForegroundColor Red
    } elseif ($Level -eq "WARNING") {
        Write-Host $logEntry -ForegroundColor Yellow
    } else {
        Write-Host $logEntry
    }
}

# Export functions
#Export-ModuleMember -Function Write-Log