# Config.ps1
# Configuration settings for the GCI-Notion365 Reporting Tool

# Check if config file exists
$configFile = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
if (Test-Path -Path $configFile) {
    # Load configuration from file
    $config = Get-Content -Path $configFile | ConvertFrom-Json
    
    # Notion API settings
    $script:notionApiKey = $config.notionApiKey
    $script:notionDatabaseId = $config.notionDatabaseId

    # Microsoft Graph API settings
    $script:tenantId = $config.tenantId
    $script:appId = $config.appId
    $script:appSecret = $config.appSecret

    # Reporting period
    $script:lookbackDays = $config.lookbackDays
} else {
    # Default settings (for development/testing)
    Write-Warning "No config file found. Using default settings. Create a config.json file for production use."
    
    # Notion API settings
    $script:notionApiKey = "secret_YourNotionAPIKeyHere"
    $script:notionDatabaseId = "YourNotionDatabaseIdHere"

    # Microsoft Graph API settings
    $script:tenantId = "YourTenantIdHere"
    $script:appId = "YourAppIdHere"
    $script:appSecret = "YourAppSecretHere"

    # Reporting period
    $script:lookbackDays = 30
}

# Calculated date values
$script:startDate = (Get-Date).AddDays(-$lookbackDays).ToString("yyyy-MM-dd")
$script:endDate = (Get-Date).ToString("yyyy-MM-dd")

# Report configuration - list of reports to collect
$script:reportsToCollect = @(
    # User activity reports
    "getOffice365ActivationsUserDetail",
    "getTeamsUserActivityUserDetail", 
    "getEmailActivityUserDetail",
    "getOneDriveActivityUserDetail",
    "getSharePointActivityUserDetail",
    "getOffice365GroupsActivityDetail",
    
    # Service usage reports
    "getTeamsDeviceUsageUserDetail",
    "getMailboxUsageDetail",
    "getOneDriveUsageAccountDetail",
    
    # Security reports
    "getEmailAppUsageUserDetail",
    "getEmailAppUsageAppsUserCounts",
    
    # License reports - special handling for these
    "subscribedSkus"
)

# Local storage for report data
$script:dataFolder = Join-Path -Path $env:ProgramData -ChildPath "GCI-Notion365\Data"
if (-not (Test-Path -Path $dataFolder)) {
    New-Item -Path $dataFolder -ItemType Directory -Force | Out-Null
}
$script:logFile = Join-Path -Path $env:ProgramData -ChildPath "GCI-Notion365\Logs\reporting.log"

# Ensure log directory exists
$logDir = Split-Path -Path $script:logFile -Parent
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

# No Export-ModuleMember is needed since we're using dot-sourcing