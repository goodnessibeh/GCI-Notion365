# M365-API-Functions.ps1
# Microsoft 365 API functions for GCI-Notion365 Reporting Tool

function Get-MsGraphToken {
    Write-Log "Acquiring Microsoft Graph token..."
    
    $tokenUrl = "https://login.microsoftonline.com/$script:tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $script:appId
        client_secret = $script:appSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    
    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        Write-Log "Token acquired successfully"
        return $response.access_token
    } catch {
        Write-Log "Failed to acquire token: $_" -Level "ERROR"
        throw
    }
}

function Get-M365Report {
    param (
        [string]$ReportName,
        [string]$AccessToken
    )
    
    Write-Log "Retrieving $ReportName report..."
    
    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept" = "application/json"
    }
    
    try {
        $reportFile = Join-Path -Path $script:dataFolder -ChildPath "$ReportName.json"
        
        # Handle special case for license information
        if ($ReportName -eq "subscribedSkus") {
            $reportUrl = "https://graph.microsoft.com/v1.0/subscribedSkus"
            $response = Invoke-RestMethod -Uri $reportUrl -Headers $headers -Method Get
            $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportFile
        }
        # Handle user activity reports
        elseif ($ReportName -like "get*") {
            $reportUrl = "https://graph.microsoft.com/v1.0/reports/$ReportName(period='D$script:lookbackDays')"
            $csvFile = Join-Path -Path $script:dataFolder -ChildPath "$ReportName.csv"
            Invoke-RestMethod -Uri $reportUrl -Headers $headers -Method Get -OutFile $csvFile
            
            # Convert CSV to structured JSON for easier processing
            $csvData = ConvertFrom-M365Report -ReportPath $csvFile
            $csvData | ConvertTo-Json | Out-File -FilePath $reportFile
        }
        # For any other reports
        else {
            $reportUrl = "https://graph.microsoft.com/v1.0/$ReportName"
            $response = Invoke-RestMethod -Uri $reportUrl -Headers $headers -Method Get
            $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportFile
        }
        
        Write-Log "Report saved to $reportFile"
        return $reportFile
    } catch {
        Write-Log "Failed to retrieve report: $_" -Level "ERROR"
        return $null
    }
}

# Export functions
#Export-ModuleMember -Function Get-MsGraphToken, Get-M365Report