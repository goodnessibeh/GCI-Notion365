# GCI-Notion365

GCI-Notion365 is a PowerShell-based solution that automatically extracts Microsoft 365 usage data and transforms it into visually appealing dashboards in Notion. Similar to paid solutions like AdminDroid, it provides comprehensive insights into Teams activity, Exchange usage, SharePoint collaboration, OneDrive file sharing, and license utilization—all without subscription fees.

**GCI-Notion365: M365 metrics delivered to Notion.**

## Overview

This PowerShell tool connects to Microsoft 365 using Graph API, extracts valuable usage metrics and statistics, then creates beautiful dashboards in Notion that give you insights into your Microsoft 365 environment. It offers features similar to paid solutions like AdminDroid but in a lightweight, customizable package.

## Features

- **Comprehensive M365 Reporting**: Collects detailed data from Teams, Exchange, SharePoint, OneDrive, and more
- **Beautiful Dashboards**: Creates visually appealing, information-rich dashboards in Notion
- **Executive Summary**: High-level metrics for management reporting
- **Manual Execution**: Run on-demand for the most up-to-date metrics
- **Error Resilient**: Robust error handling to process various Microsoft 365 report formats
- **Modular Design**: Easily extendable to add new reports and dashboards

## Dashboards

The solution creates the following dashboards in Notion:

1. **Microsoft 365 License Overview**
   - License allocation overview
   - License utilization by product
   - Detailed license usage metrics

2. **Microsoft Teams Usage Overview**
   - User engagement statistics
   - Activity breakdown
   - Top Teams users

3. **Email Usage Overview**
   - Email activity metrics
   - User engagement
   - Top email senders

4. **OneDrive & SharePoint Usage**
   - File activity metrics
   - User engagement comparison
   - External sharing statistics

5. **Microsoft 365 Executive Summary**
   - Service adoption rates
   - License utilization
   - Key metrics and recommendations

## Requirements

- Windows 10/11 or Windows Server with PowerShell 5.1+
- Microsoft 365 tenant with admin access
- Azure AD registered application with appropriate permissions
- Notion account with API access

## Prerequisites

### Microsoft Graph API Setup

1. Register an application in Microsoft Entra ID (formerly Azure AD)
   - Go to [Azure Portal](https://portal.azure.com) > Microsoft Entra ID > App registrations
   - Click "New registration"
   - Name: "GCI-Notion365 Reporting"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: Leave blank
   - Click "Register"

2. Create a client secret
   - In your application, go to "Certificates & secrets"
   - Click "New client secret"
   - Add a description and select an expiration period
   - Copy the generated secret value immediately (you won't be able to see it again)

3. Add the required API permissions
   - In your application, go to "API permissions"
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Application permissions"
   - Add the following permissions:
     - `Reports.Read.All` (For accessing usage reports)
     - `Organization.Read.All` (For organization data)
     - `Directory.Read.All` (For directory information)
     - `User.Read.All` (For user data)
     - `Group.Read.All` (For group information)
     - `TeamSettings.Read.All` (For Teams data)
     - `TeamMember.Read.All` (For Teams membership)
     - `Sites.Read.All` (For SharePoint data)
     - `Files.Read.All` (For OneDrive data)
     - `Mail.Read` (For email statistics)
     - `MailboxSettings.Read` (For mailbox settings)

4. Grant admin consent
   - After adding permissions, click "Grant admin consent for [Your Organization]"
   - You must be a global administrator to grant these permissions
   - All permissions should show as "Granted" after this step

5. Note your application details
   - Tenant ID (Directory ID): Found on the application overview page
   - Application (client) ID: Found on the application overview page
   - Client Secret: The value you copied when creating the secret

### Notion Setup

1. Create a new integration
   - Go to [Notion Integrations](https://www.notion.so/my-integrations)
   - Create a new integration
   - Enable all content capabilities

2. Create a database with the following properties:
   - ReportName (title property)
   - ReportDate (date property)
   - Various number properties for metrics (these will be created automatically)

3. Share your database with the integration
   - Open the database in Notion
   - Click "Share" and add your integration

4. Get the database ID from the URL
   - The database ID is the part of the URL after the workspace name and before any query parameters
   - Example: `https://www.notion.so/workspace/1a2b3c4d5e6f7g8h9i0j?v=...` (ID is `1a2b3c4d5e6f7g8h9i0j`)

## Project Structure

- **M365-to-Notion.ps1** - Main script
- **Config.ps1** - Configuration settings
- **Logging.ps1** - Logging functions
- **M365-API-Functions.ps1** - Microsoft 365 API interactions
- **Report-Processing.ps1** - Report data processing
- **Notion-Functions.ps1** - Notion API interactions
- **Dashboard-Creator.ps1** - Dashboard creation functions
- **Run.ps1** - Script to run the tool with direct console output

## Setup

1. Clone or download the repository
2. Create a `config.json` file in the same directory with the following structure:

```json
{
  "notionApiKey": "YOUR_NOTION_API_KEY_HERE",
  "notionDatabaseId": "YOUR_NOTION_DATABASE_ID_HERE",
  "tenantId": "YOUR_TENANT_ID_HERE",
  "appId": "YOUR_APP_ID_HERE",
  "appSecret": "YOUR_APP_SECRET_HERE",
  "lookbackDays": 30
}
```

## Usage

Run the tool by executing the Run.ps1 script:

```powershell
.\Run.ps1
```

This will:
1. Connect to Microsoft 365 via Graph API
2. Extract report data for the configured time period
3. Process and summarize the data
4. Create or update dashboards in Notion
5. Display a summary of the collected data in the console

## Troubleshooting

- **Logs**: Check logs at `%ProgramData%\GCI-Notion365\Logs`
- **Authentication Issues**: Verify API permissions and credentials in config.json
- **Notion Issues**: Ensure database is properly shared with your integration
- **CSV Processing Errors**: The tool includes robust error handling for various CSV formats
- **Permission Issues**: Ensure all required Microsoft Graph permissions are granted with admin consent

### Common Errors

- **401 Unauthorized**: Check your Microsoft Graph API credentials
- **403 Forbidden**: Ensure all required permissions are granted with admin consent
- **404 Not Found**: Verify your Notion database ID is correct and the database is shared with your integration
- **400 Bad Request**: Check the format of your API requests and parameters

## Extending the Solution

### Adding New Reports

1. Add new report names to the `$reportsToCollect` array in `Config.ps1`
2. Add processing logic for the new report in `Get-ReportSummary` function in `Report-Processing.ps1`
3. Create a new dashboard function in `Dashboard-Creator.ps1`

### Customizing Dashboards

Modify the dashboard creation functions in `Dashboard-Creator.ps1` to change:
- Chart types and data
- Table columns and data
- Visual elements and colors

## License

MIT License - See LICENSE file for details.

## License

Goodness Caleb Ibeh

## Acknowledgements

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Notion API Documentation](https://developers.notion.com/)
- Inspired by [AdminDroid](https://admindroid.com/)
