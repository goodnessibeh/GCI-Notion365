# Dashboard-Creator.ps1
# Dashboard creation functions for GCI-Notion365 Reporting Tool

function Create-AzureADLicenseDashboard {
    param (
        [string]$ReportFile
    )
    
    Write-Log "Creating Azure AD License Dashboard..."
    
    # Load license data
    $licenseData = Get-Content -Path $ReportFile | ConvertFrom-Json
    
    # Create dashboard page
    $pageId = Create-NotionDashboardPage -DatabaseId $script:notionDatabaseId -ApiKey $script:notionApiKey -Title "Microsoft 365 License Overview" -Icon "💼"
    if (-not $pageId) {
        Write-Log "Failed to create license dashboard page" -Level "ERROR"
        return
    }
    
    # Prepare license summary
    $licenseTypes = @{}
    $totalLicenses = 0
    $assignedLicenses = 0
    
    foreach ($sku in $licenseData.value) {
        $licenseName = $sku.skuPartNumber
        $totalUnits = $sku.prepaidUnits.enabled
        $consumedUnits = $sku.consumedUnits
        
        $licenseTypes[$licenseName] = @{
            "Total" = $totalUnits
            "Assigned" = $consumedUnits
            "Available" = ($totalUnits - $consumedUnits)
        }
        
        $totalLicenses += $totalUnits
        $assignedLicenses += $consumedUnits
    }
    
    # Create license overview chart
    $chartLabels = @("Assigned Licenses", "Available Licenses")
    $chartValues = @($assignedLicenses, ($totalLicenses - $assignedLicenses))
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "License Allocation Overview" -Labels $chartLabels -Values $chartValues -ChartType "pie"
    
    # Create license type breakdown chart
    $licenseLabels = $licenseTypes.Keys | Sort-Object
    $licenseAssigned = @()
    $licenseAvailable = @()
    
    foreach ($license in $licenseLabels) {
        $licenseAssigned += $licenseTypes[$license].Assigned
        $licenseAvailable += $licenseTypes[$license].Available
    }
    
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "License Allocation by Product" -Labels $licenseLabels -Values $licenseAssigned -ChartType "bar"
    
    # Create detailed license table
    $tableHeaders = @("License Name", "Total", "Assigned", "Available", "Utilization %")
    $tableRows = @()
    
    foreach ($license in $licenseTypes.Keys | Sort-Object) {
        $utilization = [math]::Round(($licenseTypes[$license].Assigned / $licenseTypes[$license].Total) * 100, 2)
        $tableRows += @($license, $licenseTypes[$license].Total, $licenseTypes[$license].Assigned, $licenseTypes[$license].Available, "$utilization%")
    }
    
    Add-NotionTableBlock -PageId $pageId -ApiKey $script:notionApiKey -TableTitle "Detailed License Allocation" -Headers $tableHeaders -Rows $tableRows
    
    Write-Log "License dashboard created successfully"
}

function Create-TeamsActivityDashboard {
    param (
        [string]$ReportFile
    )
    
    Write-Log "Creating Teams Activity Dashboard..."
    
    # Load Teams activity data
    $teamsData = Get-Content -Path $ReportFile | ConvertFrom-Json
    
    # Create dashboard page
    $pageId = Create-NotionDashboardPage -DatabaseId $script:notionDatabaseId -ApiKey $script:notionApiKey -Title "Microsoft Teams Usage Overview" -Icon "💬"
    if (-not $pageId) {
        Write-Log "Failed to create Teams dashboard page" -Level "ERROR"
        return
    }
    
    # Calculate key metrics
    $totalUsers = $teamsData.Count
    $activeUsers = ($teamsData | Where-Object { $_.TeamChatMessages -gt 0 -or $_.TeamChannelMessages -gt 0 }).Count
    $inactiveUsers = $totalUsers - $activeUsers
    $totalMessages = ($teamsData | Measure-Object -Property "TeamChatMessages" -Sum).Sum + 
                     ($teamsData | Measure-Object -Property "TeamChannelMessages" -Sum).Sum
    $totalMeetings = ($teamsData | Measure-Object -Property "MeetingCount" -Sum).Sum
    $totalCalls = ($teamsData | Measure-Object -Property "CallCount" -Sum).Sum
    
    # Create overview stats
    $headers = @{
        "Authorization" = "Bearer $script:notionApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    # Add a callout block with overview stats
    $overviewStats = "📈 **Teams Usage Summary (Last 30 Days)**
    
    • Total Users: $totalUsers
    • Active Users: $activeUsers
    • Inactive Users: $inactiveUsers
    • Total Messages Sent: $totalMessages
    • Total Meetings: $totalMeetings
    • Total Calls: $totalCalls"
    
    $statsUrl = "https://api.notion.com/v1/blocks/$pageId/children"
    $statsBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "callout"
                "callout" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $overviewStats
                            }
                        }
                    )
                    "icon" = @{
                        "emoji" = "📊"
                    }
                    "color" = "green_background"
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $statsUrl -Method Patch -Headers $headers -Body $statsBody | Out-Null
        Write-Log "Added Teams overview stats"
    } catch {
        Write-Log "Failed to add Teams overview stats: $_" -Level "ERROR"
    }
    
    # Create user activity status chart
    $activityLabels = @("Active Users", "Inactive Users")
    $activityValues = @($activeUsers, $inactiveUsers)
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "Teams User Engagement" -Labels $activityLabels -Values $activityValues -ChartType "pie"
    
    # Create activity breakdown chart
    $activityTypeLabels = @("Team Chat Messages", "Channel Messages", "Meetings", "Calls")
    $activityTypeValues = @(
        ($teamsData | Measure-Object -Property "TeamChatMessages" -Sum).Sum,
        ($teamsData | Measure-Object -Property "TeamChannelMessages" -Sum).Sum,
        $totalMeetings,
        $totalCalls
    )
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "Teams Activity Breakdown" -Labels $activityTypeLabels -Values $activityTypeValues -ChartType "bar"
    
    # Create top users table
    $topUsers = $teamsData | 
        Sort-Object -Property @{Expression={$_.TeamChatMessages + $_.TeamChannelMessages}; Descending=$true} |
        Select-Object -First 10
    
    $tableHeaders = @("User", "Team Chat Messages", "Channel Messages", "Meetings", "Calls", "Total Activities")
    $tableRows = @()
    
    foreach ($user in $topUsers) {
        $totalActivities = $user.TeamChatMessages + $user.TeamChannelMessages + $user.MeetingCount + $user.CallCount
        $tableRows += @(
            $user.UserPrincipalName, 
            $user.TeamChatMessages, 
            $user.TeamChannelMessages, 
            $user.MeetingCount, 
            $user.CallCount,
            $totalActivities
        )
    }
    
    Add-NotionTableBlock -PageId $pageId -ApiKey $script:notionApiKey -TableTitle "Top 10 Teams Users" -Headers $tableHeaders -Rows $tableRows
    
    Write-Log "Teams dashboard created successfully"
}
function Create-EmailActivityDashboard {
    param (
        [string]$ReportFile
    )
    
    Write-Log "Creating Email Activity Dashboard..."
    
    # Load email activity data
    $emailData = Get-Content -Path $ReportFile | ConvertFrom-Json
    
    # Create dashboard page
    $pageId = Create-NotionDashboardPage -DatabaseId $script:notionDatabaseId -ApiKey $script:notionApiKey -Title "Email Usage Overview" -Icon "📧"
    if (-not $pageId) {
        Write-Log "Failed to create Email dashboard page" -Level "ERROR"
        return
    }
    
    # Calculate key metrics
    $totalUsers = $emailData.Count
    $activeUsers = ($emailData | Where-Object { $_.SendCount -gt 0 -or $_.ReceiveCount -gt 0 }).Count
    $inactiveUsers = $totalUsers - $activeUsers
    $totalSent = ($emailData | Measure-Object -Property "SendCount" -Sum).Sum
    $totalReceived = ($emailData | Measure-Object -Property "ReceiveCount" -Sum).Sum
    $totalRead = ($emailData | Measure-Object -Property "ReadCount" -Sum).Sum
    
    # Create overview stats callout
    $headers = @{
        "Authorization" = "Bearer $script:notionApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    $overviewStats = "📈 **Email Usage Summary (Last 30 Days)**
    
    • Total Users: $totalUsers
    • Active Users: $activeUsers
    • Inactive Users: $inactiveUsers
    • Total Emails Sent: $totalSent
    • Total Emails Received: $totalReceived
    • Total Emails Read: $totalRead"
    
    $statsUrl = "https://api.notion.com/v1/blocks/$pageId/children"
    $statsBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "callout"
                "callout" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $overviewStats
                            }
                        }
                    )
                    "icon" = @{
                        "emoji" = "📊"
                    }
                    "color" = "blue_background"
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $statsUrl -Method Patch -Headers $headers -Body $statsBody | Out-Null
        Write-Log "Added Email overview stats"
    } catch {
        Write-Log "Failed to add Email overview stats: $_" -Level "ERROR"
    }
    
    # Create user activity status chart
    $activityLabels = @("Active Users", "Inactive Users")
    $activityValues = @($activeUsers, $inactiveUsers)
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "Email User Engagement" -Labels $activityLabels -Values $activityValues -ChartType "pie"
    
    # Create email activity chart
    $emailActivityLabels = @("Sent", "Received", "Read")
    $emailActivityValues = @($totalSent, $totalReceived, $totalRead)
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "Email Activity Overview" -Labels $emailActivityLabels -Values $emailActivityValues -ChartType "bar"
    
    # Create top email senders table
    $topSenders = $emailData | 
        Sort-Object -Property SendCount -Descending |
        Select-Object -First 10
    
    $tableHeaders = @("User", "Emails Sent", "Emails Received", "Emails Read", "Send/Receive Ratio")
    $tableRows = @()
    
    foreach ($user in $topSenders) {
        $ratio = if ($user.ReceiveCount -gt 0) { 
            [math]::Round($user.SendCount / $user.ReceiveCount, 2) 
        } else { 
            "N/A" 
        }
        
        $tableRows += @(
            $user.UserPrincipalName, 
            $user.SendCount, 
            $user.ReceiveCount, 
            $user.ReadCount,
            $ratio
        )
    }
    
    Add-NotionTableBlock -PageId $pageId -ApiKey $script:notionApiKey -TableTitle "Top 10 Email Senders" -Headers $tableHeaders -Rows $tableRows
    
    Write-Log "Email dashboard created successfully"
}

function Create-OneDriveSharePointDashboard {
    param (
        [string]$OneDriveReportFile,
        [string]$SharePointReportFile
    )
    
    Write-Log "Creating OneDrive & SharePoint Dashboard..."
    
    # Load report data
    $oneDriveData = Get-Content -Path $OneDriveReportFile | ConvertFrom-Json
    $sharePointData = Get-Content -Path $SharePointReportFile | ConvertFrom-Json
    
    # Create dashboard page
    $pageId = Create-NotionDashboardPage -DatabaseId $script:notionDatabaseId -ApiKey $script:notionApiKey -Title "OneDrive & SharePoint Usage" -Icon "📁"
    if (-not $pageId) {
        Write-Log "Failed to create OneDrive & SharePoint dashboard page" -Level "ERROR"
        return
    }
    
    # Calculate OneDrive metrics
    $totalOneDriveUsers = $oneDriveData.Count
    $activeOneDriveUsers = ($oneDriveData | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count
    $totalOneDriveFiles = ($oneDriveData | Measure-Object -Property "ViewedOrEditedFileCount" -Sum).Sum
    $totalOneDriveSharedExt = ($oneDriveData | Measure-Object -Property "SharedExternallyFileCount" -Sum).Sum
    $totalOneDriveSharedInt = ($oneDriveData | Measure-Object -Property "SharedInternallyFileCount" -Sum).Sum
    $totalOneDriveSynced = ($oneDriveData | Measure-Object -Property "SyncedFileCount" -Sum).Sum
    
    # Calculate SharePoint metrics
    $totalSPUsers = $sharePointData.Count
    $activeSPUsers = ($sharePointData | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count
    $totalSPFiles = ($sharePointData | Measure-Object -Property "ViewedOrEditedFileCount" -Sum).Sum
    $totalSPSharedExt = ($sharePointData | Measure-Object -Property "SharedExternallyFileCount" -Sum).Sum
    $totalSPSharedInt = ($sharePointData | Measure-Object -Property "SharedInternallyFileCount" -Sum).Sum
    
    # Create overview stats callout
    $headers = @{
        "Authorization" = "Bearer $script:notionApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    $overviewStats = "📈 **OneDrive & SharePoint Usage Summary (Last 30 Days)**
    
    **OneDrive Stats:**
    • Active Users: $activeOneDriveUsers of $totalOneDriveUsers
    • Files Viewed/Edited: $totalOneDriveFiles
    • Files Shared Externally: $totalOneDriveSharedExt
    • Files Shared Internally: $totalOneDriveSharedInt
    • Files Synced: $totalOneDriveSynced
    
    **SharePoint Stats:**
    • Active Users: $activeSPUsers of $totalSPUsers
    • Files Viewed/Edited: $totalSPFiles
    • Files Shared Externally: $totalSPSharedExt
    • Files Shared Internally: $totalSPSharedInt"
    
    $statsUrl = "https://api.notion.com/v1/blocks/$pageId/children"
    $statsBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "callout"
                "callout" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $overviewStats
                            }
                        }
                    )
                    "icon" = @{
                        "emoji" = "📊"
                    }
                    "color" = "purple_background"
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    try {
        Invoke-RestMethod -Uri $statsUrl -Method Patch -Headers $headers -Body $statsBody | Out-Null
        Write-Log "Added OneDrive & SharePoint overview stats"
    } catch {
        Write-Log "Failed to add OneDrive & SharePoint overview stats: $_" -Level "ERROR"
    }
    
    # Create user activity comparison chart
    $activityLabels = @("OneDrive Active Users", "OneDrive Inactive Users", "SharePoint Active Users", "SharePoint Inactive Users")
    $activityValues = @(
        $activeOneDriveUsers, 
        ($totalOneDriveUsers - $activeOneDriveUsers),
        $activeSPUsers,
        ($totalSPUsers - $activeSPUsers)
    )
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "User Engagement Comparison" -Labels $activityLabels -Values $activityValues -ChartType "bar"
    
    # Create file activity comparison chart
    $fileActivityLabels = @(
        "OneDrive - Viewed/Edited", 
        "OneDrive - Shared Internally", 
        "OneDrive - Shared Externally",
        "SharePoint - Viewed/Edited",
        "SharePoint - Shared Internally",
        "SharePoint - Shared Externally"
    )
    $fileActivityValues = @(
        $totalOneDriveFiles,
        $totalOneDriveSharedInt,
        $totalOneDriveSharedExt,
        $totalSPFiles,
        $totalSPSharedInt,
        $totalSPSharedExt
    )
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "File Activity Breakdown" -Labels $fileActivityLabels -Values $fileActivityValues -ChartType "bar"
    
    # Create top OneDrive users table
    $topOneDriveUsers = $oneDriveData | 
        Sort-Object -Property ViewedOrEditedFileCount -Descending |
        Select-Object -First 10
    
    $tableHeaders = @("User", "Files Viewed/Edited", "Files Shared Internally", "Files Shared Externally", "Files Synced")
    $tableRows = @()
    
    foreach ($user in $topOneDriveUsers) {
        $tableRows += @(
            $user.UserPrincipalName, 
            $user.ViewedOrEditedFileCount, 
            $user.SharedInternallyFileCount, 
            $user.SharedExternallyFileCount,
            $user.SyncedFileCount
        )
    }
    
    Add-NotionTableBlock -PageId $pageId -ApiKey $script:notionApiKey -TableTitle "Top 10 OneDrive Users" -Headers $tableHeaders -Rows $tableRows
    
    Write-Log "OneDrive & SharePoint dashboard created successfully"
}

function Create-ExecutiveSummaryDashboard {
    param (
        [hashtable]$ReportSummaries
    )
    
    Write-Log "Creating Executive Summary Dashboard..."
    
    # Create dashboard page
    $pageId = Create-NotionDashboardPage -DatabaseId $script:notionDatabaseId -ApiKey $script:notionApiKey -Title "Microsoft 365 Executive Summary" -Icon "🏢"
    if (-not $pageId) {
        Write-Log "Failed to create Executive Summary dashboard page" -Level "ERROR"
        return
    }
    
    $headers = @{
        "Authorization" = "Bearer $script:notionApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    # Extract key metrics from report summaries
    $licenseData = $ReportSummaries["subscribedSkus"]
    $teamsData = $ReportSummaries["getTeamsUserActivityUserDetail"]
    $emailData = $ReportSummaries["getEmailActivityUserDetail"]
    $oneDriveData = $ReportSummaries["getOneDriveActivityUserDetail"]
    $sharePointData = $ReportSummaries["getSharePointActivityUserDetail"]
    
    # Create introduction block
    $intro = "# Microsoft 365 Executive Summary
    
This dashboard provides a high-level overview of your Microsoft 365 environment as of $(Get-Date -Format "MMMM d, yyyy"). Use this summary to quickly assess the health and usage patterns across your Microsoft 365 services.

## Key Highlights

* **License Utilization**: $($licenseData["Assigned Licenses"])/$($licenseData["Total Licenses"]) licenses assigned ($(($licenseData["Assigned Licenses"]/$licenseData["Total Licenses"]*100).ToString("F1"))%)
* **Active Users**: Teams ($($teamsData["Active Users"])), Email ($($emailData["Active Users"])), OneDrive ($($oneDriveData["Active Users"]))
* **Team Collaboration**: $($teamsData["Meetings Attended"]) meetings and $($teamsData["Total Messages"]) messages exchanged
* **Email Volume**: $($emailData["Emails Sent"]) emails sent and $($emailData["Emails Received"]) received
* **File Activity**: $($oneDriveData["Files Viewed/Edited"] + $sharePointData["Files Viewed/Edited"]) files accessed across OneDrive and SharePoint
    "
    
    $introUrl = "https://api.notion.com/v1/blocks/$pageId/children"
    $introBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "paragraph"
                "paragraph" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $intro
                            }
                        }
                    )
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $introUrl -Method Patch -Headers $headers -Body $introBody | Out-Null
        Write-Log "Added Executive Summary introduction"
    } catch {
        Write-Log "Failed to add Executive Summary introduction: $_" -Level "ERROR"
    }
    
    # Create service health overview chart
    $serviceLabels = @("Teams", "Email", "OneDrive", "SharePoint")
    $serviceValues = @(
        ($teamsData["Active Users"] / $teamsData["Total Users"] * 100),
        ($emailData["Active Users"] / $emailData["Total Users"] * 100),
        ($oneDriveData["Active Users"] / $oneDriveData["Total Users"] * 100),
        ($sharePointData["Active Users"] / $sharePointData["Total Users"] * 100)
    )
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "Service Adoption Rates (%)" -Labels $serviceLabels -Values $serviceValues -ChartType "bar"
    
    # Create license overview chart
    $licenseLabels = @("Assigned Licenses", "Available Licenses")
    $licenseValues = @(
        $licenseData["Assigned Licenses"],
        ($licenseData["Total Licenses"] - $licenseData["Assigned Licenses"])
    )
    Create-NotionChartBlock -PageId $pageId -ApiKey $script:notionApiKey -ChartTitle "License Utilization" -Labels $licenseLabels -Values $licenseValues -ChartType "pie"
    
    # Create productivity metrics table
    $tableHeaders = @("Service", "Active Users", "Total Activity", "Key Metrics")
    $tableRows = @(
        @(
            "Microsoft Teams", 
            "$($teamsData["Active Users"]) of $($teamsData["Total Users"])", 
            "$($teamsData["Total Messages"]) messages", 
            "$($teamsData["Meetings Attended"]) meetings, $($teamsData["Calls Participated"]) calls"
        ),
        @(
            "Exchange Online", 
            "$($emailData["Active Users"]) of $($emailData["Total Users"])", 
            "$($emailData["Emails Sent"] + $emailData["Emails Received"]) emails", 
            "$($emailData["Emails Sent"]) sent, $($emailData["Emails Received"]) received"
        ),
        @(
            "OneDrive", 
            "$($oneDriveData["Active Users"]) of $($oneDriveData["Total Users"])", 
            "$($oneDriveData["Files Viewed/Edited"]) files accessed", 
            "$($oneDriveData["Files Shared Internally"] + $oneDriveData["Files Shared Externally"]) files shared"
        ),
        @(
            "SharePoint", 
            "$($sharePointData["Active Users"]) of $($sharePointData["Total Users"])", 
            "$($sharePointData["Files Viewed/Edited"]) files accessed", 
            "$($sharePointData["Files Shared Internally"] + $sharePointData["Files Shared Externally"]) files shared"
        )
    )
    
    Add-NotionTableBlock -PageId $pageId -ApiKey $script:notionApiKey -TableTitle "Service Usage Summary" -Headers $tableHeaders -Rows $tableRows
    
    # Add recommendations section
    $recommendations = "## Recommendations

Based on the current usage data, consider the following actions:

* **License Optimization**: Review unused licenses to potentially reduce costs
* **Teams Adoption**: Encourage more use of Teams meetings for the $(($teamsData["Total Users"] - $teamsData["Active Users"])) inactive users
* **Security Review**: Monitor external sharing activities across OneDrive ($($oneDriveData["Files Shared Externally"]) files) and SharePoint ($($sharePointData["Files Shared Externally"]) files)
* **Training Opportunities**: Provide additional training for low-adoption services"
    
    $recUrl = "https://api.notion.com/v1/blocks/$pageId/children"
    $recBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "paragraph"
                "paragraph" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $recommendations
                            }
                        }
                    )
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $recUrl -Method Patch -Headers $headers -Body $recBody | Out-Null
        Write-Log "Added Executive Summary recommendations"
    } catch {
        Write-Log "Failed to add Executive Summary recommendations: $_" -Level "ERROR"
    }
    
    Write-Log "Executive Summary dashboard created successfully"
}

# Export functions
#Export-ModuleMember -Function Create-AzureADLicenseDashboard, Create-TeamsActivityDashboard, Create-EmailActivityDashboard, Create-OneDriveSharePointDashboard, Create-ExecutiveSummaryDashboard