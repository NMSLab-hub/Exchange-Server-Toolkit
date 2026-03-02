<#
.SYNOPSIS
    Full Exchange On-Prem Configuration Report (Optimized for Large Environments)
.DESCRIPTION
    - Mailbox summary only
    - Extra Mailboxes details only
    - Transport Rules, Connectors, Domains, Policies, Compliance info
    - Collapsible sections for readability
    - Safe for production (read-only)
	Prepared by Nur Mohammod Tohad
#>

# -----------------------
# Setup
# -----------------------
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$Date       = Get-Date -Format "yyyy-MM-dd_HH-mm"
$OutputFile = Join-Path $ScriptDir "Exchange_FullReport_$Date.html"

function Safe-Run {
    param([scriptblock]$Command, [string]$Name="")
    try { & $Command } catch { Write-Warning "Could not run $Name : $_"; return $null }
}

# -----------------------
# Organization & Exchange Info
# -----------------------
$OrgName     = Safe-Run { (Get-OrganizationConfig).Name } -Name "Get-OrganizationConfig"
$ExVer       = Safe-Run { (Get-ExchangeServer | Select-Object -First 1).AdminDisplayVersion } -Name "Get-ExchangeServer"
$GeneratedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# -----------------------
# Health Check Data
# -----------------------
$ExchangeServers = Safe-Run { Get-ExchangeServer } -Name "Get-ExchangeServer"
$DatabaseStatus = Safe-Run { Get-MailboxDatabase -Status } -Name "Get-MailboxDatabase"
$ReplicationStatus = Safe-Run { Get-MailboxDatabaseCopyStatus } -Name "Get-MailboxDatabaseCopyStatus"
$QueueStatus = Safe-Run { Get-Queue } -Name "Get-Queue"

# Health Check Summary
$HealthSummary = [PSCustomObject]@{
    TotalServers         = ($ExchangeServers | Measure-Object).Count
    DatabasesHealthy     = ($DatabaseStatus | Where-Object { $_.Mounted -eq $true }).Count
    DatabasesTotal       = ($DatabaseStatus | Measure-Object).Count
    ReplicationIssues    = ($ReplicationStatus | Where-Object { $_.Status -notin @("Mounted", "Healthy") }).Count
    QueuesWithIssues     = ($QueueStatus | Where-Object { $_.MessageCount -gt 50 }).Count
    DiskSpaceIssues      = ($ExchangeServers | Where-Object { 
        $disk = Safe-Run { Get-WmiObject Win32_LogicalDisk -ComputerName $_.Name -Filter "DeviceID='C:'" } -Name "Get-DiskSpace"
        if ($disk) { ($disk.FreeSpace / $disk.Size) * 100 -lt 20 } else { $false }
    }).Count
}

# -----------------------
# Mailbox Summary, Arbitration, & Discovery Mailboxes
# -----------------------
$AllMailboxes   = Safe-Run { Get-Mailbox -ResultSize Unlimited } -Name "Get-Mailbox"
$OrgCount       = ($AllMailboxes | Measure-Object).Count
$FilteredMail   = $AllMailboxes | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox') }
$RealCount      = ($FilteredMail | Measure-Object).Count
$ArbitrationMailboxes = Safe-Run { Get-Mailbox -Arbitration -ResultSize Unlimited } -Name "Get-Mailbox Arbitration"
$OtherSystemMailboxes = Safe-Run { Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -notin @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','ArbitrationMailbox') } } -Name "Get-Mailbox OtherSystem"

# -----------------------
# Transport / Mail Flow Rules
# -----------------------
$TransportRules = Safe-Run { Get-TransportRule | Sort-Object Priority } -Name "Get-TransportRule"
$TransportRulesFull = @()
if ($TransportRules) {
    $idx = 0
    foreach ($r in $TransportRules) {
        $idx++
        $full = Safe-Run { Get-TransportRule -Identity $r.Identity | Select-Object * } -Name "Get-TransportRule Detail"
        $applied = @()
        foreach ($prop in @('SentTo','SentToScope','From','FromScope','SentToAddresses','FromAddresses')) {
            if ($full.PSObject.Properties.Match($prop)) {
                $val = $full.PSObject.Properties[$prop].Value
                if ($val -is [array]) { $val = $val -join "; " }
                $applied += "$prop : $val"
            } else { $applied += "$prop : " }
        }
        $TransportRulesFull += [PSCustomObject]@{
            Idx         = $idx
            Name        = $r.Name
            Identity    = $r.Identity
            Priority    = $r.Priority
            Mode        = $r.Mode
            State       = $r.State
            Comments    = $r.Comments
            Conditions  = ($full.Conditions -as [array]) -join "`n"
            Actions     = ($full.Actions -as [array]) -join "`n"
            Exceptions  = ($full.Exceptions -as [array]) -join "`n"
            Applied     = ($applied -join "`n")
            FullConfig  = ($full | Format-List * | Out-String)
        }
    }
}

# -----------------------
# Other Configuration
# -----------------------
$SendConnectors       = Safe-Run { Get-SendConnector } -Name "Get-SendConnector"
$ReceiveConnectors    = Safe-Run { Get-ReceiveConnector } -Name "Get-ReceiveConnector"
$AcceptedDomains      = Safe-Run { Get-AcceptedDomain } -Name "Get-AcceptedDomain"
$EmailAddressPolicies = Safe-Run { Get-EmailAddressPolicy } -Name "Get-EmailAddressPolicy"
$JournalRules         = Safe-Run { Get-JournalRule } -Name "Get-JournalRule"
$RetentionPolicies    = Safe-Run { Get-RetentionPolicy } -Name "Get-RetentionPolicy"
$RetentionTags        = Safe-Run { Get-RetentionPolicyTag } -Name "Get-RetentionPolicyTag"
$DlpPolicies          = Safe-Run { Get-DlpPolicy -ErrorAction SilentlyContinue } -Name "Get-DlpPolicy"

# -----------------------
# HTML / CSS / JS
# -----------------------
$css = @"
<style>
body {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
    margin: 0;
    padding: 24px;
    background: #F5F6F5;
    color: #212121;
    line-height: 1.6;
    overflow: auto;
}
.container {
    width: 90%;
    max-width: 1400px;
    min-width: 320px;
    margin: auto;
    background: #FFFFFF;
    border-radius: 12px;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
    padding: 24px;
}
.h1 {
    font-size: 30px;
    font-weight: 600;
    color: #003087;
    margin-bottom: 8px;
}
.subtitle {
    color: #616161;
    font-size: 14px;
    margin-bottom: 24px;
}
.section {
    margin-bottom: 20px;
    border-radius: 8px;
    overflow: hidden;
}
.section-header {
    background: #003087;
    color: #FFFFFF;
    padding: 12px 16px;
    font-size: 18px;
    font-weight: 600;
    cursor: pointer;
    display: flex;
    justify-content: space-between;
    align-items: center;
    transition: background 0.2s ease;
}
.section-header:hover {
    background: #00215E;
}
.section-header.health-check {
    border-left: 4px solid #2E7D32;
    font-weight: 700;
}
.section-content {
    padding: 16px;
    background: #FFFFFF;
    transition: all 0.3s ease;
    overflow-x: auto;
    overflow-y: auto;
    max-height: 600px;
}
.table {
    width: 100%;
    min-width: 600px;
    border-collapse: collapse;
    font-size: 13px;
}
.table th {
    background: #005A9E;
    color: #FFFFFF;
    padding: 12px;
    text-align: left;
    font-weight: 600;
}
.table td {
    padding: 12px;
    border-bottom: 1px solid #B0BEC5;
    vertical-align: top;
}
.prebox {
    background: #F5F6F5;
    color: #212121;
    padding: 12px;
    border-radius: 6px;
    font-family: 'Consolas', monospace;
    font-size: 12px;
    overflow: auto;
    white-space: pre-wrap;
}
.badge-good {
    background: #2E7D32;
    color: #FFFFFF;
    padding: 4px 8px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
}
.badge-bad {
    background: #D32F2F;
    color: #FFFFFF;
    padding: 4px 8px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
}
.badge-warning {
    background: #F57C00;
    color: #FFFFFF;
    padding: 4px 8px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
}
.toggle-btn {
    background: #005A9E;
    color: #FFFFFF;
    border: none;
    padding: 8px 16px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 12px;
    transition: background 0.2s ease;
}
.toggle-btn:hover {
    background: #003087;
}
.hidden {
    display: none;
}
@keyframes slideDown {
    from { max-height: 0; opacity: 0; }
    to { max-height: 1000px; opacity: 1; }
}
.section-content.hidden {
    animation: slideDown 0.3s ease forwards;
}
@media screen and (max-width: 768px) {
    .container {
        width: 95%;
        padding: 16px;
    }
    .h1 {
        font-size: 24px;
    }
    .section-header {
        font-size: 16px;
    }
    .table {
        font-size: 12px;
    }
}
</style>
<script>
function toggle(id) {
    var element = document.getElementById(id);
    if (element) {
        element.classList.toggle('hidden');
    }
}
</script>
"@

# -----------------------
# HTML Start
# -----------------------
$Html = @"
<html>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Exchange Server Health Check Report</title>
<link href='https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&family=Consolas&display=swap' rel='stylesheet'>
$css
</head>
<body>
<div class='container'>
<h1 class='h1'>Exchange Server Configuration Detailed Report</h1>
<p class='subtitle'>Org: $OrgName | Exchange: $ExVer | Generated: $GeneratedOn</p>
"@

# -----------------------
# Health Check Section
# -----------------------
$Html += @"
<div class='section'>
    <div class='section-header health-check' onclick="toggle('health-check')">
        <span>Exchange Health Check</span>
        <span></span>
    </div>
    <div class='section-content' id='health-check'>
        <table class='table'>
            <thead>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Total Exchange Servers</td>
                    <td>$($HealthSummary.TotalServers)</td>
                    <td>$(if ($HealthSummary.TotalServers -gt 0) { "<span class='badge-good'>Operational</span>" } else { "<span class='badge-bad'>No Servers</span>" })</td>
                </tr>
                <tr>
                    <td>Healthy Databases</td>
                    <td>$($HealthSummary.DatabasesHealthy)/$($HealthSummary.DatabasesTotal)</td>
                    <td>$(if ($HealthSummary.DatabasesHealthy -eq $HealthSummary.DatabasesTotal -and $HealthSummary.DatabasesTotal -gt 0) { "<span class='badge-good'>All Healthy</span>" } else { "<span class='badge-warning'>Issues Detected</span>" })</td>
                </tr>
                <tr>
                    <td>Replication Issues</td>
                    <td>$($HealthSummary.ReplicationIssues)</td>
                    <td>$(if ($HealthSummary.ReplicationIssues -eq 0) { "<span class='badge-good'>No Issues</span>" } else { "<span class='badge-bad'>Issues Detected</span>" })</td>
                </tr>
                <tr>
                    <td>Mail Queues with Issues</td>
                    <td>$($HealthSummary.QueuesWithIssues)</td>
                    <td>$(if ($HealthSummary.QueuesWithIssues -eq 0) { "<span class='badge-good'>No Issues</span>" } else { "<span class='badge-warning'>Issues Detected</span>" })</td>
                </tr>
                <tr>
                    <td>Servers with Low Disk Space</td>
                    <td>$($HealthSummary.DiskSpaceIssues)</td>
                    <td>$(if ($HealthSummary.DiskSpaceIssues -eq 0) { "<span class='badge-good'>No Issues</span>" } else { "<span class='badge-bad'>Issues Detected</span>" })</td>
                </tr>
            </tbody>
        </table>
    </div>
</div>
"@

# -----------------------
# Mailbox Summary
# -----------------------
$Html += @"
<div class='section'>
    <div class='section-header' onclick="toggle('mailbox-summary')">
        <span>Mailbox Summary</span>
        <span></span>
    </div>
    <div class='section-content' id='mailbox-summary'>
        <table class='table'>
            <thead>
                <tr>
                    <th>Total Mailboxes</th>
                    <th>Filtered Mailboxes</th>
                    <th>Discovery Mailboxes</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>$OrgCount</td>
                    <td>$RealCount</td>
                    <td>$($OtherSystemMailboxes.Count)</td>
                </tr>
            </tbody>
        </table>
        <p style='font-size: 13px; margin-top: 8px;'>Note: Arbitration Mailboxes Count: $($ArbitrationMailboxes.Count)</p>
    </div>
</div>
"@

# -----------------------
# Discovery System Mailboxes Details
# -----------------------
if ($OtherSystemMailboxes.Count -gt 0) {
    $Html += @"
    <div class='section'>
        <div class='section-header' onclick="toggle('discovery-mailboxes')">
            <span>Discovery System Mailboxes Details</span>
            <span></span>
        </div>
        <div class='section-content' id='discovery-mailboxes'>
            <table class='table'>
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>DisplayName</th>
                        <th>Type</th>
                        <th>Primary SMTP</th>
                        <th>Database</th>
                        <th>Server</th>
                    </tr>
                </thead>
                <tbody>
"@
    foreach ($mb in $OtherSystemMailboxes) {
        $Html += "<tr>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.Name) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.DisplayName) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.RecipientTypeDetails) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.PrimarySmtpAddress) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.Database) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.ServerName) )</td>
        </tr>"
    }
    $Html += @"
                </tbody>
            </table>
        </div>
    </div>
"@
}

# -----------------------
# Arbitration Mailboxes Details
# -----------------------
if ($ArbitrationMailboxes.Count -gt 0) {
    $Html += @"
    <div class='section'>
        <div class='section-header' onclick="toggle('arbitration-mailboxes')">
            <span>Arbitration Mailboxes Details</span>
            <span></span>
        </div>
        <div class='section-content' id='arbitration-mailboxes'>
            <table class='table'>
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>DisplayName</th>
                        <th>Type</th>
                        <th>Primary SMTP</th>
                        <th>Database</th>
                        <th>Server</th>
                    </tr>
                </thead>
                <tbody>
"@
    foreach ($mb in $ArbitrationMailboxes) {
        $Html += "<tr>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.Name) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.DisplayName) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.RecipientTypeDetails) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.PrimarySmtpAddress) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.Database) )</td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($mb.ServerName) )</td>
        </tr>"
    }
    $Html += @"
                </tbody>
            </table>
        </div>
    </div>
"@
}

# -----------------------
# Transport Rules Section
# -----------------------
if ($TransportRulesFull.Count -gt 0) {
    $Html += @"
    <div class='section'>
        <div class='section-header' onclick="toggle('transport-rules')">
            <span>Transport / Mail Flow Rules</span>
            <span></span>
        </div>
        <div class='section-content' id='transport-rules'>
            <table class='table'>
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Priority</th>
                        <th>Status</th>
                        <th>Applied Accounts</th>
                        <th>Conditions</th>
                        <th>Actions</th>
                        <th>Exceptions</th>
                    </tr>
                </thead>
                <tbody>
"@
    foreach ($r in $TransportRulesFull) {
        $statusBadge = if ($r.State -eq "Enabled") { "<span class='badge-good'>Enabled</span>" } else { "<span class='badge-bad'>Disabled</span>" }
        $Html += "<tr>
            <td><strong>$( [System.Web.HttpUtility]::HtmlEncode($r.Name) )</strong><br/><small>Identity: $( [System.Web.HttpUtility]::HtmlEncode($r.Identity) )</small></td>
            <td>$( [System.Web.HttpUtility]::HtmlEncode($r.Priority) )</td>
            <td>$statusBadge</td>
            <td><pre class='prebox'>$( [System.Web.HttpUtility]::HtmlEncode($r.Applied) )</pre></td>
            <td><pre class='prebox'>$( [System.Web.HttpUtility]::HtmlEncode($r.Conditions) )</pre></td>
            <td><pre class='prebox'>$( [System.Web.HttpUtility]::HtmlEncode($r.Actions) )</pre></td>
            <td><pre class='prebox'>$( [System.Web.HttpUtility]::HtmlEncode($r.Exceptions) )</pre></td>
        </tr>"
        $Html += "<tr><td colspan='7'><button class='toggle-btn' onclick=""toggle('rule_full_$($r.Idx)')"" aria-expanded='false'>Toggle Full Config</button><div id='rule_full_$($r.Idx)' class='hidden'><pre class='prebox'>$( [System.Web.HttpUtility]::HtmlEncode($r.FullConfig) )</pre></div></td></tr>"
    }
    $Html += "</tbody></table></div></div>"
}

# -----------------------
# Generic Function to Add Configuration Sections
# -----------------------
function Add-ConfigSection($Title, $Data, $SectionId) {
    if ($Data -and $Data.Count -gt 0) {
        $HtmlSection = "<div class='section'><div class='section-header' onclick=""toggle('$SectionId')""><span>$Title</span><span></span></div><div class='section-content' id='$SectionId'><table class='table'><thead><tr>"
        $props = $Data[0].PSObject.Properties | ForEach-Object { $_.Name }
        foreach ($p in $props) { $HtmlSection += "<th>$( [System.Web.HttpUtility]::HtmlEncode($p) )</th>" }
        $HtmlSection += "</tr></thead><tbody>"
        foreach ($item in $Data) {
            $HtmlSection += "<tr>"
            foreach ($p in $props) {
                $val = $item.PSObject.Properties[$p].Value
                if ($val -is [array]) { $val = $val -join "; " }
                $HtmlSection += "<td>$( [System.Web.HttpUtility]::HtmlEncode($val) )</td>"
            }
            $HtmlSection += "</tr>"
        }
        $HtmlSection += "</tbody></table></div></div>"
        return $HtmlSection
    }
    return ""
}

# -----------------------
# Add Remaining Sections
# -----------------------
$Html += Add-ConfigSection "Send Connectors" $SendConnectors "send-connectors"
$Html += Add-ConfigSection "Receive Connectors" $ReceiveConnectors "receive-connectors"
$Html += Add-ConfigSection "Accepted Domains" $AcceptedDomains "accepted-domains"
$Html += Add-ConfigSection "Email Address Policies" $EmailAddressPolicies "email-policies"
$Html += Add-ConfigSection "Journal Rules" $JournalRules "journal-rules"
$Html += Add-ConfigSection "Retention Policies" $RetentionPolicies "retention-policies"
$Html += Add-ConfigSection "Retention Tags" $RetentionTags "retention-tags"
if ($DlpPolicies) { $Html += Add-ConfigSection "DLP Policies" $DlpPolicies "dlp-policies" }

# -----------------------
# Finish HTML & Save
# -----------------------
$Html += "</div></body></html>"
$Html | Out-File -FilePath $OutputFile -Encoding UTF8
Write-Host "Exchange Health Check report generated: $OutputFile" -ForegroundColor Cyan
Start-Process $OutputFile