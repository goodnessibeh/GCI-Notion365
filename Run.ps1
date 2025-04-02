# Run.ps1
# Script to run GCI-Notion365 tool with direct console output

# Enable detailed error information
$ErrorActionPreference = "Continue"
$ProgressPreference = "Continue"
$VerbosePreference = "Continue"

# Script path information
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$MainScript = Join-Path -Path $ScriptPath -ChildPath "M365-to-Notion.ps1"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  GCI-Notion365: M365 metrics delivered to Notion" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "Running script: $MainScript" -ForegroundColor Cyan
Write-Host "Current directory: $ScriptPath" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Check if main script exists
if (-not (Test-Path -Path $MainScript)) {
    Write-Host "ERROR: Main script not found at: $MainScript" -ForegroundColor Red
    Write-Host "Please ensure the script files are properly installed." -ForegroundColor Red
    exit 1
}

# Check if config file exists
$ConfigFile = Join-Path -Path $ScriptPath -ChildPath "config.json"
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Host "ERROR: Configuration file not found at: $ConfigFile" -ForegroundColor Red
    Write-Host "Please ensure config.json is present." -ForegroundColor Red
    exit 1
}

# Check config file content
try {
    $configContent = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    
    # Check for default values
    if ($configContent.notionApiKey -eq "YOUR_NOTION_API_KEY_HERE" -or
        $configContent.tenantId -eq "YOUR_TENANT_ID_HERE") {
        Write-Host "WARNING: Configuration file appears to contain default values." -ForegroundColor Yellow
        Write-Host "Please edit config.json with your actual API keys and settings." -ForegroundColor Yellow
        
        $proceed = Read-Host "Do you want to proceed anyway? (y/n)"
        if ($proceed.ToLower() -ne "y") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    }
} catch {
    Write-Host "ERROR: Failed to read configuration file: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Starting main process..." -ForegroundColor Green
Write-Host ""

try {
    # Run each individual file directly
    Write-Host "Loading configuration..." -ForegroundColor Cyan
    . "$ScriptPath\Config.ps1"
    
    Write-Host "Loading logging functions..." -ForegroundColor Cyan
    . "$ScriptPath\Logging.ps1"
    
    Write-Host "Loading API functions..." -ForegroundColor Cyan
    . "$ScriptPath\M365-API-Functions.ps1"
    
    Write-Host "Loading report processing functions..." -ForegroundColor Cyan
    . "$ScriptPath\Report-Processing.ps1"
    
    Write-Host "Loading Notion functions..." -ForegroundColor Cyan
    . "$ScriptPath\Notion-Functions.ps1"
    
    Write-Host "Loading dashboard creator functions..." -ForegroundColor Cyan
    . "$ScriptPath\Dashboard-Creator.ps1"
    
    Write-Host "Executing main script logic..." -ForegroundColor Green
    
    # Run the main script logic - instead of the file itself
    # Get Microsoft Graph access token
    $accessToken = Get-MsGraphToken
    
    # Create hashtable to store all report summaries
    $reportSummaries = @{}
    
    foreach ($reportName in $reportsToCollect) {
        # Get the report from Microsoft 365
        $reportFile = Get-M365Report -ReportName $reportName -AccessToken $accessToken
        
        if ($reportFile) {
            if ($reportName -eq "subscribedSkus") {
                # Special handling for license data (JSON format)
                $reportData = Get-Content -Path $reportFile | ConvertFrom-Json
                
                # Generate license summary with proper nested property handling
                $totalLicenses = 0
                $assignedLicenses = 0
                
                foreach ($sku in $reportData.value) {
                    if ($sku.prepaidUnits -and $sku.prepaidUnits.enabled) {
                        $totalLicenses += $sku.prepaidUnits.enabled
                    }
                    if ($sku.consumedUnits) {
                        $assignedLicenses += $sku.consumedUnits
                    }
                }
                
                $summary = @{
                    "Total Licenses" = $totalLicenses
                    "Assigned Licenses" = $assignedLicenses
                }
                
                $reportSummaries[$reportName] = $summary
            } else {
                # Process the report data (CSV format)
                $csvFile = Join-Path -Path $dataFolder -ChildPath "$reportName.csv"
                if (Test-Path -Path $csvFile) {
                    $reportData = ConvertFrom-M365Report -ReportPath $csvFile
                
                    if ($reportData) {
                        # Generate summary metrics
                        $summary = Get-ReportSummary -ReportData $reportData -ReportName $reportName
                        $reportSummaries[$reportName] = $summary
                        
                        # Keep Notion updates (they will error but remain in place for later)
                        try {
                            Update-NotionDatabase -DatabaseId $notionDatabaseId -ApiKey $notionApiKey -ReportName $reportName -Summary $summary -ReportDate (Get-Date)
                        } catch {
                            Write-Host "Notion update failed for $reportName (will be fixed later): $_" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "Warning: CSV file not found for $reportName" -ForegroundColor Yellow
                }
            }
        }
    }
    
    Write-Host "Creating AdminDroid-style dashboards in Notion..." -ForegroundColor Cyan
    
    # Keep dashboard creation (they will error but remain in place for later)
    try {
        # Create Azure AD License Dashboard
        $licenseReportFile = Join-Path -Path $dataFolder -ChildPath "subscribedSkus.json"
        if (Test-Path $licenseReportFile) {
            Create-AzureADLicenseDashboard -ReportFile $licenseReportFile
        }
        
        # Create Teams Activity Dashboard
        $teamsReportFile = Join-Path -Path $dataFolder -ChildPath "getTeamsUserActivityUserDetail.json"
        if (Test-Path $teamsReportFile) {
            Create-TeamsActivityDashboard -ReportFile $teamsReportFile
        }
        
        # Create Email Activity Dashboard
        $emailReportFile = Join-Path -Path $dataFolder -ChildPath "getEmailActivityUserDetail.json"
        if (Test-Path $emailReportFile) {
            Create-EmailActivityDashboard -ReportFile $emailReportFile
        }
        
        # Create OneDrive & SharePoint Dashboard
        $oneDriveReportFile = Join-Path -Path $dataFolder -ChildPath "getOneDriveActivityUserDetail.json"
        $sharePointReportFile = Join-Path -Path $dataFolder -ChildPath "getSharePointActivityUserDetail.json"
        if (Test-Path $oneDriveReportFile -and Test-Path $sharePointReportFile) {
            Create-OneDriveSharePointDashboard -OneDriveReportFile $oneDriveReportFile -SharePointReportFile $sharePointReportFile
        }
        
        # Create Executive Summary Dashboard
        Create-ExecutiveSummaryDashboard -ReportSummaries $reportSummaries
    } catch {
        Write-Host "Notion dashboard creation failed (will be fixed later): $_" -ForegroundColor Yellow
    }
    
    Write-Host "Data collection completed." -ForegroundColor Green
    
    # Display summary of collected data
    Write-Host ""
    Write-Host "Summary of Collected Data:" -ForegroundColor Green
    foreach ($reportName in $reportSummaries.Keys) {
        Write-Host ""
        Write-Host "$reportName :" -ForegroundColor Cyan
        foreach ($metric in $reportSummaries[$reportName].Keys) {
            Write-Host "  $metric`: $($reportSummaries[$reportName][$metric])" -ForegroundColor White
        }
    }
    
    $exitCode = 0
} catch {
    Write-Host "ERROR: An unexpected error occurred:" -ForegroundColor Red
    Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    
    $exitCode = 1
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Execution Complete" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Ended at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan

if ($exitCode -eq 0) {
    Write-Host "Status: SUCCESS" -ForegroundColor Green
    Write-Host "NOTE: Notion integration errors are expected until database is configured." -ForegroundColor Yellow
} else {
    Write-Host "Status: FAILED" -ForegroundColor Red
}

# Use Read-Host instead of ReadKey for better compatibility
Write-Host "Press Enter to exit..." -ForegroundColor Cyan
Read-Host
exit $exitCode