<#
.SYNOPSIS
    Generates an HTML report of all mailboxes and their locations (Exchange Online vs On-Premises)

.DESCRIPTION
    This script connects to Exchange Online and generates a comprehensive HTML report
    showing all mailboxes (User, Room, Shared, etc.) and whether they are located
    in Exchange Online or still on the on-premises Exchange server.

.PARAMETER OutputPath
    The path where the HTML report will be saved. Default is current directory.

.EXAMPLE
    .\Generate-MailboxLocationReport.ps1
    .\Generate-MailboxLocationReport.ps1 -OutputPath "C:\Reports\"
#>

param(
    [string]$OutputPath = "."
)

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "ExchangeOnlineManagement module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to import ExchangeOnlineManagement module. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    exit 1
}

# Connect to Exchange Online
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowProgress $true
    Write-Host "Connected to Exchange Online successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit 1
}

# Get all mailboxes with detailed information
Write-Host "Retrieving mailbox information..." -ForegroundColor Yellow

try {
    # Get all Exchange Online mailboxes
    $onlineMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object `
        DisplayName, 
        PrimarySmtpAddress, 
        RecipientTypeDetails, 
        Database,
        ServerName,
        ExchangeVersion,
        RemoteRecipientType,
        @{Name="MailboxLocation"; Expression={ "Exchange Online" }},
        @{Name="DatabaseLocation"; Expression={
            if ($_.Database) { $_.Database } else { "N/A" }
        }},
        WhenCreated,
        LastLogonTime

    # Try to get remote mailboxes (on-premises mailboxes in hybrid environment)
    $remoteMailboxes = @()
    try {
        $remoteMailboxes = Get-RemoteMailbox -ResultSize Unlimited | Select-Object `
            DisplayName, 
            PrimarySmtpAddress, 
            RecipientTypeDetails, 
            @{Name="Database"; Expression={ "On-Premises Exchange" }},
            @{Name="ServerName"; Expression={ "On-Premises Server" }},
            ExchangeVersion,
            RemoteRecipientType,
            @{Name="MailboxLocation"; Expression={ "On-Premises" }},
            @{Name="DatabaseLocation"; Expression={ "On-Premises Exchange" }},
            WhenCreated,
            @{Name="LastLogonTime"; Expression={ $null }}
        
        Write-Host "Found $($remoteMailboxes.Count) remote (on-premises) mailboxes" -ForegroundColor Green
    }
    catch {
        Write-Host "No remote mailboxes found or not in hybrid environment" -ForegroundColor Yellow
    }

    # Combine all mailboxes
    $mailboxes = @($onlineMailboxes) + @($remoteMailboxes)
    
    Write-Host "Found $($onlineMailboxes.Count) Exchange Online mailboxes" -ForegroundColor Green
    Write-Host "Total mailboxes: $($mailboxes.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve mailbox information: $($_.Exception.Message)"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Generate HTML report
$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$fileName = "MailboxLocationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$fullPath = Join-Path $OutputPath $fileName

# Group mailboxes by location for summary
$locationSummary = $mailboxes | Group-Object MailboxLocation | Sort-Object Name

# Create HTML content
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Mailbox Location Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 20px; border-radius: 5px; }
        .summary { background-color: #f5f5f5; padding: 15px; margin: 20px 0; border-radius: 5px; }
        .summary-item { display: inline-block; margin: 10px 20px 10px 0; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .online { background-color: #d4edda; }
        .onprem { background-color: #fff3cd; }
        .unknown { background-color: #f8d7da; }
        .filter-buttons { margin: 20px 0; }
        .filter-btn { padding: 8px 16px; margin: 5px; border: none; border-radius: 3px; cursor: pointer; }
        .filter-btn.active { background-color: #0078d4; color: white; }
        .filter-btn:not(.active) { background-color: #e9ecef; }
    </style>
    <script>
        function filterTable(location) {
            var table = document.getElementById("mailboxTable");
            var rows = table.getElementsByTagName("tr");
            var buttons = document.getElementsByClassName("filter-btn");
            
            // Reset button styles
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].classList.remove("active");
            }
            
            // Set active button
            event.target.classList.add("active");
            
            for (var i = 1; i < rows.length; i++) {
                var locationCell = rows[i].getElementsByTagName("td")[3];
                if (location === "all" || locationCell.innerHTML === location) {
                    rows[i].style.display = "";
                } else {
                    rows[i].style.display = "none";
                }
            }
        }
    </script>
</head>
<body>
    <div class="header">
        <h1>Exchange Mailbox Location Report</h1>
        <p>Generated on: $reportDate</p>
    </div>
    
    <div class="summary">
        <h2>Summary</h2>
"@

# Add summary statistics
foreach ($group in $locationSummary) {
    $htmlContent += "<div class='summary-item'><strong>$($group.Name):</strong> $($group.Count) mailboxes</div>"
}

$htmlContent += @"
        <div class='summary-item'><strong>Total:</strong> $($mailboxes.Count) mailboxes</div>
    </div>
    
    <div class="filter-buttons">
        <h3>Filter by Location:</h3>
        <button class="filter-btn active" onclick="filterTable('all')">All</button>
        <button class="filter-btn" onclick="filterTable('Exchange Online')">Exchange Online</button>
        <button class="filter-btn" onclick="filterTable('On-Premises')">On-Premises</button>
        <button class="filter-btn" onclick="filterTable('Unknown')">Unknown</button>
    </div>
    
    <table id="mailboxTable">
        <thead>
            <tr>
                <th>Display Name</th>
                <th>Email Address</th>
                <th>Mailbox Type</th>
                <th>Location</th>
                <th>Database</th>
                <th>Server Name</th>
                <th>Created Date</th>
                <th>Last Logon</th>
            </tr>
        </thead>
        <tbody>
"@

# Add mailbox data
foreach ($mailbox in $mailboxes | Sort-Object DisplayName) {
    $rowClass = switch ($mailbox.MailboxLocation) {
        "Exchange Online" { "online" }
        "On-Premises" { "onprem" }
        default { "unknown" }
    }
    
    $lastLogon = if ($mailbox.LastLogonTime) { $mailbox.LastLogonTime.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
    $created = if ($mailbox.WhenCreated) { $mailbox.WhenCreated.ToString("yyyy-MM-dd") } else { "Unknown" }
    
    $htmlContent += @"
            <tr class="$rowClass">
                <td>$($mailbox.DisplayName)</td>
                <td>$($mailbox.PrimarySmtpAddress)</td>
                <td>$($mailbox.RecipientTypeDetails)</td>
                <td>$($mailbox.MailboxLocation)</td>
                <td>$($mailbox.DatabaseLocation)</td>
                <td>$($mailbox.ServerName)</td>
                <td>$created</td>
                <td>$lastLogon</td>
            </tr>
"@
}

$htmlContent += @"
        </tbody>
    </table>
    
    <div style="margin-top: 30px; padding: 15px; background-color: #f8f9fa; border-radius: 5px;">
        <h3>Legend:</h3>
        <div style="margin: 5px 0;"><span style="background-color: #d4edda; padding: 2px 8px; border-radius: 3px;">Exchange Online</span> - Mailboxes hosted in Microsoft 365</div>
        <div style="margin: 5px 0;"><span style="background-color: #fff3cd; padding: 2px 8px; border-radius: 3px;">On-Premises</span> - Mailboxes hosted on local Exchange servers</div>
        <div style="margin: 5px 0;"><span style="background-color: #f8d7da; padding: 2px 8px; border-radius: 3px;">Unknown</span> - Location could not be determined</div>
    </div>
</body>
</html>
"@

# Save the HTML report
try {
    $htmlContent | Out-File -FilePath $fullPath -Encoding UTF8
    Write-Host "Report generated successfully: $fullPath" -ForegroundColor Green
    
    # Display summary
    Write-Host "`nSummary:" -ForegroundColor Cyan
    foreach ($group in $locationSummary) {
        Write-Host "  $($group.Name): $($group.Count) mailboxes" -ForegroundColor White
    }
    Write-Host "  Total: $($mailboxes.Count) mailboxes" -ForegroundColor White
}
catch {
    Write-Error "Failed to save report: $($_.Exception.Message)"
}
finally {
    # Disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected successfully" -ForegroundColor Green
}

Write-Host "`nScript completed. Report saved to: $fullPath" -ForegroundColor Green
