# M365-to-Notion.ps1
# Main script for GCI-Notion365 Reporting Tool

# Import required modules
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Dot-source all the module scripts
. "$ScriptPath\Config.ps1"
. "$ScriptPath\Logging.ps1"
. "$ScriptPath\M365-API-Functions.ps1"
. "$ScriptPath\Report-Processing.ps1"
. "$ScriptPath\Notion-Functions.ps1"
. "$ScriptPath\Dashboard-Creator.ps1"

# Main execution
Write-Log "Starting M365 to Notion reporting process"

try {
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
                # Generate license summary
                $summary = @{
                    "Total Licenses" = ($reportData.value | Measure-Object -Property "prepaidUnits.enabled" -Sum).Sum
                    "Assigned Licenses" = ($reportData.value | Measure-Object -Property "consumedUnits" -Sum).Sum
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
                        
                        # Update Notion with the report data
                        Update-NotionDatabase -DatabaseId $notionDatabaseId -ApiKey $notionApiKey -ReportName $reportName -Summary $summary -ReportDate (Get-Date)
                    }
                } else {
                    Write-Log "CSV file not found for $reportName" -Level "WARNING"
                }
            }
        }
    }
    
    # Create AdminDroid-style dashboards
    Write-Log "Creating AdminDroid-style dashboards in Notion..."
    
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
    
    Write-Log "M365 to Notion reporting process completed successfully"
} catch {
    Write-Log "An error occurred during the reporting process: $_" -Level "ERROR"
    exit 1
}