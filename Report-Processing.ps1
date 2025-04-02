# Report-Processing.ps1
# Report processing functions for the GCI-Notion365 Reporting Tool

function ConvertFrom-M365Report {
    param (
        [string]$ReportPath
    )
    
    Write-Log "Processing report: $ReportPath"
    
    if (-not (Test-Path -Path $ReportPath)) {
        Write-Log "Report file not found: $ReportPath" -Level "ERROR"
        return $null
    }
    
    # Determine report type for specialized handling
    $reportType = "generic"
    if ($ReportPath -match "TeamsUserActivityUserDetail|TeamsDeviceUsageUserDetail|MailboxUsageDetail") {
        $reportType = "teams"
    }
    elseif ($ReportPath -match "OneDriveActivityUserDetail|SharePointActivityUserDetail") {
        $reportType = "onedrive"
    }
    elseif ($ReportPath -match "EmailAppUsageUserDetail") {
        $reportType = "email"
    }
    
    try {
        # Use simpler approach with Import-Csv for all report types
        Write-Log "Using simple CSV import for $reportType report"
        
        # Create a hashtable to track headers that have been seen
        $seenHeaders = @{}
        
        # M365 reports have a special first line with metadata, we need to skip it
        $csvContent = Get-Content -Path $ReportPath -Encoding UTF8 | Select-Object -Skip 1
        
        # Extract just the header line
        $headerLine = $csvContent[0]
        $originalHeaders = $headerLine -split ',' | ForEach-Object { $_.Trim('"') }
        
        # Process headers to make them unique
        $uniqueHeaders = @()
        foreach ($header in $originalHeaders) {
            # Clean up header name
            $cleanHeader = $header -replace '[^a-zA-Z0-9_]', '_'
            
            # Handle date headers
            if ($cleanHeader -match '^\d{4}-\d{2}-\d{2}$') {
                $cleanHeader = "Date_$cleanHeader"
            }
            
            # Handle numeric headers
            if ($cleanHeader -match '^\d+$') {
                $cleanHeader = "Value_$cleanHeader"
            }
            
            # Handle empty headers
            if ([string]::IsNullOrWhiteSpace($cleanHeader)) {
                $cleanHeader = "Column_$($uniqueHeaders.Count)"
            }
            
            # Ensure uniqueness by adding counter if needed
            $baseHeader = $cleanHeader
            $counter = 1
            while ($seenHeaders.ContainsKey($cleanHeader)) {
                $cleanHeader = "${baseHeader}_$counter"
                $counter++
            }
            
            $seenHeaders[$cleanHeader] = $true
            $uniqueHeaders += $cleanHeader
        }
        
        # Create a new temporary CSV with fixed headers
        $tempFile = "$ReportPath.processed.csv"
        $uniqueHeaderLine = $uniqueHeaders -join ','
        $uniqueHeaderLine | Out-File -FilePath $tempFile -Encoding utf8
        $csvContent | Select-Object -Skip 1 | Out-File -FilePath $tempFile -Encoding utf8 -Append
        
        # Use Import-Csv to parse the fixed CSV
        $parsedData = Import-Csv -Path $tempFile -ErrorAction Stop
        
        # Clean up temporary file
        if (Test-Path -Path $tempFile) {
            Remove-Item -Path $tempFile -Force
        }
        
        Write-Log "Successfully processed report with $($parsedData.Count) records"
        return $parsedData
    } 
    catch {
        Write-Log "Standard processing failed: $_" -Level "ERROR"
        
        # Use more robust CSV parsing as fallback
        try {
            Write-Log "Attempting direct processing..." -Level "WARNING"
            
            # Last resort - manually create objects from raw data
            $rawContent = Get-Content -Path $ReportPath -Raw -Encoding UTF8
            $contentLines = $rawContent -split "\r?\n" | Where-Object { $_.Trim() -ne "" }
            
            # Skip metadata line
            $contentLines = $contentLines | Select-Object -Skip 1
            
            # If we have content, process it line by line
            if ($contentLines.Count -gt 1) {
                $headerLine = $contentLines[0]
                $headers = $headerLine -split ',' | ForEach-Object { 
                    $_.Trim('"') -replace '[^a-zA-Z0-9_]', '_' 
                }
                
                # Ensure unique headers
                $uniqueHeaders = @()
                $seenHeaders = @{}
                
                foreach ($header in $headers) {
                    if ([string]::IsNullOrWhiteSpace($header)) {
                        $header = "Column_$($uniqueHeaders.Count)"
                    }
                    
                    # Ensure uniqueness by adding a counter suffix if needed
                    $baseHeader = $header
                    $counter = 1
                    while ($seenHeaders.ContainsKey($header)) {
                        $header = "${baseHeader}_$counter"
                        $counter++
                    }
                    
                    $seenHeaders[$header] = $true
                    $uniqueHeaders += $header
                }
                
                # Process data lines
                $objects = @()
                for ($i = 1; $i -lt $contentLines.Count; $i++) {
                    $dataLine = $contentLines[$i]
                    $values = $dataLine -split ',' | ForEach-Object { $_.Trim('"') }
                    
                    # Create object with properties
                    $obj = New-Object PSObject
                    for ($j = 0; $j -lt [Math]::Min($uniqueHeaders.Count, $values.Count); $j++) {
                        # Add property and value with Force to handle potential duplicates
                        Add-Member -InputObject $obj -MemberType NoteProperty -Name $uniqueHeaders[$j] -Value $values[$j] -Force
                    }
                    
                    $objects += $obj
                }
                
                Write-Log "Successfully processed report with $($objects.Count) records using direct method"
                return $objects
            } else {
                # Return empty array with metadata if no content lines
                Write-Log "No data rows found in report" -Level "WARNING"
                return @(
                    [PSCustomObject]@{
                        ReportName = $ReportPath
                        IsEmpty = $true
                    }
                )
            }
        }
        catch {
            Write-Log "All processing methods failed: $_" -Level "ERROR"
            
            # Return minimal valid data to prevent script failure
            return @(
                [PSCustomObject]@{
                    ReportName = $ReportPath
                    ProcessingError = $true
                    ErrorMessage = $_.ToString()
                }
            )
        }
    }
}

function Get-ReportSummary {
    param (
        [array]$ReportData,
        [string]$ReportName
    )
    
    Write-Log "Generating summary for $ReportName"
    
    # Check if we have valid data
    if ($null -eq $ReportData -or $ReportData.Count -eq 0) {
        return @{
            "Total Records" = 0
            "Status" = "No data available"
        }
    }
    
    # Check if we have error data from fallback method
    if ($ReportData[0].PSObject.Properties.Name -contains "ProcessingError") {
        return @{
            "Total Records" = 0
            "Status" = "Processing error occurred"
            "Error" = $ReportData[0].ErrorMessage
        }
    }
    
    # Create appropriate summaries based on report type
    switch -Wildcard ($ReportName) {
        "*Activations*" {
            # First check the actual column names in the data
            $columnNames = ($ReportData | Get-Member -MemberType NoteProperty).Name
            $windowsColumn = $columnNames | Where-Object { $_ -match "Windows|windows" } | Select-Object -First 1
            $macColumn = $columnNames | Where-Object { $_ -match "Mac|mac" } | Select-Object -First 1
            $mobileColumn = $columnNames | Where-Object { $_ -match "Mobile|mobile|iOS|Android" } | Select-Object -First 1
            $webColumn = $columnNames | Where-Object { $_ -match "Web|web|browser" } | Select-Object -First 1
            
            $summary = @{
                "Total Users" = $ReportData.Count
                "Windows Activations" = if ($windowsColumn) { ($ReportData | Where-Object { $_.$windowsColumn -gt 0 }).Count } else { 0 }
                "Mac Activations" = if ($macColumn) { ($ReportData | Where-Object { $_.$macColumn -gt 0 }).Count } else { 0 }
                "Mobile Activations" = if ($mobileColumn) { ($ReportData | Where-Object { $_.$mobileColumn -gt 0 }).Count } else { 0 }
                "Web Activations" = if ($webColumn) { ($ReportData | Where-Object { $_.$webColumn -gt 0 }).Count } else { 0 }
            }
        }
        "*Teams*" {
            # Check for Teams-specific columns
            $columnNames = ($ReportData | Get-Member -MemberType NoteProperty).Name
            $chatColumn = $columnNames | Where-Object { $_ -match "Chat|chat|TeamChat" } | Select-Object -First 1
            $channelColumn = $columnNames | Where-Object { $_ -match "Channel|channel" } | Select-Object -First 1
            $meetingColumn = $columnNames | Where-Object { $_ -match "Meeting|meeting" } | Select-Object -First 1
            $callColumn = $columnNames | Where-Object { $_ -match "Call|call" } | Select-Object -First 1
            
            # Safely measure with error handling
            $chatCount = if ($chatColumn) { 
                try { ($ReportData | Measure-Object -Property $chatColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $channelCount = if ($channelColumn) { 
                try { ($ReportData | Measure-Object -Property $channelColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $meetingCount = if ($meetingColumn) { 
                try { ($ReportData | Measure-Object -Property $meetingColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $callCount = if ($callColumn) { 
                try { ($ReportData | Measure-Object -Property $callColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $activeUsers = if ($chatColumn -or $channelColumn) { 
                try {
                    ($ReportData | Where-Object { 
                        ($chatColumn -and $_.$chatColumn -gt 0) -or 
                        ($channelColumn -and $_.$channelColumn -gt 0) 
                    }).Count 
                } catch { 0 }
            } else { 0 }
            
            $summary = @{
                "Total Users" = $ReportData.Count
                "Active Users" = $activeUsers
                "Total Messages" = $chatCount + $channelCount
                "Calls Participated" = $callCount
                "Meetings Attended" = $meetingCount
            }
        }
        "*Email*" {
            # Check for Email-specific columns
            $columnNames = ($ReportData | Get-Member -MemberType NoteProperty).Name
            $sendColumn = $columnNames | Where-Object { $_ -match "Send|send|Sent|sent" } | Select-Object -First 1
            $receiveColumn = $columnNames | Where-Object { $_ -match "Receive|receive|Received|received" } | Select-Object -First 1
            $readColumn = $columnNames | Where-Object { $_ -match "Read|read" } | Select-Object -First 1
            
            $sendCount = if ($sendColumn) { 
                try { ($ReportData | Measure-Object -Property $sendColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $receiveCount = if ($receiveColumn) { 
                try { ($ReportData | Measure-Object -Property $receiveColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $readCount = if ($readColumn) { 
                try { ($ReportData | Measure-Object -Property $readColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $activeUsers = if ($sendColumn -or $receiveColumn) {
                try {
                    ($ReportData | Where-Object { 
                        ($sendColumn -and $_.$sendColumn -gt 0) -or 
                        ($receiveColumn -and $_.$receiveColumn -gt 0) 
                    }).Count
                } catch { 0 }
            } else { 0 }
            
            $summary = @{
                "Total Users" = $ReportData.Count
                "Active Users" = $activeUsers
                "Emails Sent" = $sendCount
                "Emails Received" = $receiveCount
                "Emails Read" = $readCount
            }
        }
        "*OneDrive*" {
            # Check for OneDrive-specific columns
            $columnNames = ($ReportData | Get-Member -MemberType NoteProperty).Name
            $viewedColumn = $columnNames | Where-Object { $_ -match "View|view|Edit|edit|File|file" } | Select-Object -First 1
            $syncedColumn = $columnNames | Where-Object { $_ -match "Sync|sync" } | Select-Object -First 1
            $sharedIntColumn = $columnNames | Where-Object { $_ -match "Internal|internal|Inside|inside" } | Select-Object -First 1
            $sharedExtColumn = $columnNames | Where-Object { $_ -match "External|external|Outside|outside" } | Select-Object -First 1
            
            $viewedCount = if ($viewedColumn) { 
                try { ($ReportData | Measure-Object -Property $viewedColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $syncedCount = if ($syncedColumn) { 
                try { ($ReportData | Measure-Object -Property $syncedColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $sharedIntCount = if ($sharedIntColumn) { 
                try { ($ReportData | Measure-Object -Property $sharedIntColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $sharedExtCount = if ($sharedExtColumn) { 
                try { ($ReportData | Measure-Object -Property $sharedExtColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $activeUsers = if ($viewedColumn) {
                try { ($ReportData | Where-Object { $_.$viewedColumn -gt 0 }).Count } catch { 0 }
            } else { 0 }
            
            $summary = @{
                "Total Users" = $ReportData.Count
                "Active Users" = $activeUsers
                "Files Viewed/Edited" = $viewedCount
                "Files Synced" = $syncedCount
                "Files Shared Internally" = $sharedIntCount
                "Files Shared Externally" = $sharedExtCount
            }
        }
        "*SharePoint*" {
            # Check for SharePoint-specific columns
            $columnNames = ($ReportData | Get-Member -MemberType NoteProperty).Name
            $viewedColumn = $columnNames | Where-Object { $_ -match "View|view|Edit|edit|File|file" } | Select-Object -First 1
            $sharedIntColumn = $columnNames | Where-Object { $_ -match "Internal|internal|Inside|inside" } | Select-Object -First 1
            $sharedExtColumn = $columnNames | Where-Object { $_ -match "External|external|Outside|outside" } | Select-Object -First 1
            
            $viewedCount = if ($viewedColumn) { 
                try { ($ReportData | Measure-Object -Property $viewedColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $sharedIntCount = if ($sharedIntColumn) { 
                try { ($ReportData | Measure-Object -Property $sharedIntColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $sharedExtCount = if ($sharedExtColumn) { 
                try { ($ReportData | Measure-Object -Property $sharedExtColumn -Sum).Sum } catch { 0 } 
            } else { 0 }
            
            $activeUsers = if ($viewedColumn) {
                try { ($ReportData | Where-Object { $_.$viewedColumn -gt 0 }).Count } catch { 0 }
            } else { 0 }
            
            $summary = @{
                "Total Users" = $ReportData.Count
                "Active Users" = $activeUsers
                "Files Viewed/Edited" = $viewedCount
                "Files Shared Internally" = $sharedIntCount
                "Files Shared Externally" = $sharedExtCount
            }
        }
        "subscribedSkus" {
            # Handle the special case of license information separately
            $totalLicenses = 0
            $assignedLicenses = 0
            
            if ($ReportData.value) {
                foreach ($sku in $ReportData.value) {
                    if ($sku.prepaidUnits -and $sku.prepaidUnits.enabled) {
                        $totalLicenses += $sku.prepaidUnits.enabled
                    }
                    if ($sku.consumedUnits) {
                        $assignedLicenses += $sku.consumedUnits
                    }
                }
            }
            
            $summary = @{
                "Total Licenses" = $totalLicenses
                "Assigned Licenses" = $assignedLicenses
            }
        }
        default {
            # Generic summary for any other report type
            $summary = @{
                "Total Records" = $ReportData.Count
                "Column Names" = ($ReportData | Get-Member -MemberType NoteProperty).Name -join ", "
            }
        }
    }
    
    return $summary
}

# No Export-ModuleMember needed since we're using dot-sourcing