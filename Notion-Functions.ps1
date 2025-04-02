# Notion-Functions.ps1
# Notion API functions for GCI-Notion365 Reporting Tool

function Update-NotionDatabase {
    param (
        [string]$DatabaseId,
        [string]$ApiKey,
        [string]$ReportName,
        [hashtable]$Summary,
        [datetime]$ReportDate
    )
    
    Write-Log "Updating Notion database for $ReportName"
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    # Check if entry already exists
    $queryUrl = "https://api.notion.com/v1/databases/$DatabaseId/query"
    $queryBody = @{
        filter = @{
            property = "ReportName"
            rich_text = @{
                equals = $ReportName
            }
        }
    } | ConvertTo-Json -Depth 10
    
    try {
        $existingPages = Invoke-RestMethod -Uri $queryUrl -Method Post -Headers $headers -Body $queryBody
        $existingPage = $existingPages.results | Where-Object { 
            (Get-Date $_.properties.ReportDate.date.start) -ge (Get-Date).AddDays(-2) 
        } | Select-Object -First 1
    } catch {
        Write-Log "Error querying Notion database: $_" -Level "ERROR"
        $existingPage = $null
    }
    
    # Prepare properties for Notion page
    $properties = @{
        "ReportName" = @{
            "title" = @(
                @{
                    "text" = @{
                        "content" = $ReportName
                    }
                }
            )
        }
        "ReportDate" = @{
            "date" = @{
                "start" = $ReportDate.ToString("yyyy-MM-dd")
            }
        }
    }
    
    # Add summary metrics as properties
    foreach ($key in $Summary.Keys) {
        $properties[$key] = @{
            "number" = $Summary[$key]
        }
    }
    
    if ($existingPage) {
        # Update existing page
        $pageId = $existingPage.id
        $updateUrl = "https://api.notion.com/v1/pages/$pageId"
        
        $updateBody = @{
            "properties" = $properties
        } | ConvertTo-Json -Depth 10
        
        try {
            Invoke-RestMethod -Uri $updateUrl -Method Patch -Headers $headers -Body $updateBody | Out-Null
            Write-Log "Updated existing Notion page for $ReportName"
        } catch {
            Write-Log "Failed to update Notion page: $_" -Level "ERROR"
        }
    } else {
        # Create new page
        $createUrl = "https://api.notion.com/v1/pages"
        
        $createBody = @{
            "parent" = @{
                "database_id" = $DatabaseId
            }
            "properties" = $properties
        } | ConvertTo-Json -Depth 10
        
        try {
            Invoke-RestMethod -Uri $createUrl -Method Post -Headers $headers -Body $createBody | Out-Null
            Write-Log "Created new Notion page for $ReportName"
        } catch {
            Write-Log "Failed to create Notion page: $_" -Level "ERROR"
        }
    }
}

function Create-NotionDashboardPage {
    param (
        [string]$DatabaseId,
        [string]$ApiKey,
        [string]$Title,
        [string]$Icon = "📊"
    )
    
    Write-Log "Creating Notion dashboard page: $Title"
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    $createUrl = "https://api.notion.com/v1/pages"
    
    $createBody = @{
        "parent" = @{
            "database_id" = $DatabaseId
        }
        "icon" = @{
            "emoji" = $Icon
        }
        "properties" = @{
            "title" = @{
                "title" = @(
                    @{
                        "text" = @{
                            "content" = $Title
                        }
                    }
                )
            }
            "Date" = @{
                "date" = @{
                    "start" = (Get-Date).ToString("yyyy-MM-dd")
                }
            }
        }
    } | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $createUrl -Method Post -Headers $headers -Body $createBody
        Write-Log "Created new Notion dashboard page with ID: $($response.id)"
        return $response.id
    } catch {
        Write-Log "Failed to create Notion dashboard page: $_" -Level "ERROR"
        return $null
    }
}

function Create-NotionChartBlock {
    param(
        [string]$PageId,
        [string]$ApiKey,
        [string]$ChartTitle,
        [array]$Labels,
        [array]$Values,
        [string]$ChartType = "bar" # bar, line, pie
    )
    
    Write-Log "Adding $ChartType chart to Notion page: $ChartTitle"
    
    # For security, we need to create a CSV in the format needed for Notion charts
    $csvData = @()
    for ($i = 0; $i -lt $Labels.Count; $i++) {
        $csvData += [PSCustomObject]@{
            "Category" = $Labels[$i]
            "Value" = $Values[$i]
        }
    }
    
    $csvString = $csvData | ConvertTo-Csv | Out-String
    
    # Notion API endpoint for blocks
    $url = "https://api.notion.com/v1/blocks/$PageId/children"
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    # Create heading for chart
    $headingBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "heading_2"
                "heading_2" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $ChartTitle
                            }
                        }
                    )
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $headingBody | Out-Null
        Write-Log "Added heading for chart: $ChartTitle"
    } catch {
        Write-Log "Failed to add heading for chart: $_" -Level "ERROR"
    }
    
    # Workaround for chart blocks since Notion API doesn't directly support chart creation
    # We'll create a callout block with instructions to manually create a chart
    $chartInstructions = "👉 Chart data prepared for '$ChartTitle'. To visualize: 
    1. Create a new $ChartType chart block below
    2. Select 'Paste CSV' option
    3. Paste this data:
    
    $csvString"
    
    $chartBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "callout"
                "callout" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $chartInstructions
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
        Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $chartBody | Out-Null
        Write-Log "Added chart data block for: $ChartTitle"
        return $true
    } catch {
        Write-Log "Failed to add chart data block: $_" -Level "ERROR"
        return $false
    }
}

function Add-NotionTableBlock {
    param (
        [string]$PageId,
        [string]$ApiKey,
        [string]$TableTitle,
        [array]$Headers,
        [array]$Rows
    )
    
    Write-Log "Adding table to Notion page: $TableTitle"
    
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }
    
    # Notion API endpoint for blocks
    $url = "https://api.notion.com/v1/blocks/$PageId/children"
    
    # Create heading for table
    $headingBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "heading_2"
                "heading_2" = @{
                    "rich_text" = @(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $TableTitle
                            }
                        }
                    )
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $headingBody | Out-Null
        Write-Log "Added heading for table: $TableTitle"
    } catch {
        Write-Log "Failed to add heading for table: $_" -Level "ERROR"
    }
    
    # Create table block
    $tableRows = @(
        # Header row
        @{
            "cells" = @(
                foreach ($header in $Headers) {
                    ,@(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = $header
                                "link" = $null
                            }
                            "annotations" = @{
                                "bold" = $true
                            }
                        }
                    )
                }
            )
        }
    )
    
    # Data rows
    foreach ($row in $Rows) {
        $tableRow = @{
            "cells" = @(
                foreach ($cell in $row) {
                    ,@(
                        @{
                            "type" = "text"
                            "text" = @{
                                "content" = "$cell"
                                "link" = $null
                            }
                        }
                    )
                }
            )
        }
        $tableRows += $tableRow
    }
    
    $tableBody = @{
        "children" = @(
            @{
                "object" = "block"
                "type" = "table"
                "table" = @{
                    "table_width" = $Headers.Count
                    "has_column_header" = $true
                    "has_row_header" = $false
                    "children" = $tableRows
                }
            }
        )
    } | ConvertTo-Json -Depth 10 -Compress
    
    try {
        Invoke-RestMethod -Uri $url -Method Patch -Headers $headers -Body $tableBody | Out-Null
        Write-Log "Added table: $TableTitle"
        return $true
    } catch {
        Write-Log "Failed to add table: $_" -Level "ERROR"
        return $false
    }
}

# Export functions
#Export-ModuleMember -Function Update-NotionDatabase, Create-NotionDashboardPage, Create-NotionChartBlock, Add-NotionTableBlock