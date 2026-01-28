<#
.SYNOPSIS
    Generates and emails an Entra ID User Sign-In Report.

.DESCRIPTION
    Reads Entra ID users and their sign-in activity from Microsoft Graph, filters to users who have NOT signed in
    within the last 90 days (or never signed in), exports the result to CSV, and emails the report.
    
    Additionally, the script excludes any users that are direct members of specific "resource mailbox" groups
    (shared mailboxes, room mailboxes, scheduling mailboxes) so they never appear in the report.

    Intended for Azure Automation runbooks using Managed Identity and Microsoft Graph.

.NOTES
    Author:       Peter Busted
    Created:      12-01-2026
    Last Edited:  28-01-2026
    Version:      1.3

    PLATFORM:
    - This script is intended to run as an Azure Automation Runbook.
    - Uses Managed Identity for authentication.
    - Runtime environment: PowerShell 7.2 or higher.
    
    REQUIRED MODULES:
    - Microsoft.Graph.Authentication 2.25.0
    - Microsoft.Graph.Users 2.25.0
    - Microsoft.Graph.Users.Actions 2.25.0
    - Microsoft.Graph.Groups 2.25.0

    REQUIRED PERMISSIONS (Managed Identity / App permissions in Graph):
    - User.Read.All            (Read user objects)
    - AuditLog.Read.All        (Read sign-in activity)
    - GroupMember.Read.All     (Read group memberships for exclusion)
    - Mail.Send                (Send email with the report)
#>

# ---------------------------------------------------------
# Connect to Microsoft Graph (Managed Identity)
# ---------------------------------------------------------
# Azure Automation uses the Automation Account Managed Identity to authenticate to Graph.
Connect-MgGraph -Identity

# ---------------------------------------------------------
# Build exclusion list (direct group members only)
# ---------------------------------------------------------
# Purpose:
# We do not want resource mailbox accounts in the sign-in report.
# These accounts are grouped in:
#   - Resources-SharedMailboxes
#   - Resources-RoomMailboxes
#   - Resources-SchedulingMailboxes
#
# Since these groups do not contain nested groups, we keep it simple:
# 1) Find each group by display name
# 2) Read direct members
# 3) Store user IDs in a HashSet for fast lookups when filtering $Users

$ExcludeGroupNames = @(
    "Resources-SharedMailboxes",
    "Resources-RoomMailboxes",
    "Resources-SchedulingMailboxes"
)

# HashSet gives fast "contains" checks and avoids duplicates automatically
$ExcludedUserIds = [System.Collections.Generic.HashSet[string]]::new()

foreach ($GroupName in $ExcludeGroupNames) {

    # Look up the group ID using its display name
    $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -Property Id, DisplayName | Select-Object -First 1

    if (-not $Group) {
        # If a group is missing, we warn and continue (script still runs)
        Write-Warning "Exclude group not found: $GroupName"
        continue
    }

    Write-Output "Loading direct members for exclude group: $($Group.DisplayName)"

    # Get direct members of the group (not transitive)
    $Members = Get-MgGroupMember -GroupId $Group.Id -All

    foreach ($m in $Members) {

        # Group membership can include non-user objects (devices, service principals, groups)
        # We only want to exclude real users.
        $odataType = $null
        if ($m.AdditionalProperties -and $m.AdditionalProperties.ContainsKey("@odata.type")) {
            $odataType = $m.AdditionalProperties["@odata.type"]
        }

        if ($odataType -eq "#microsoft.graph.user") {
            [void]$ExcludedUserIds.Add([string]$m.Id)
        }
    }
}

Write-Output "Excluded direct user count (union of all groups): $($ExcludedUserIds.Count)"

# ---------------------------------------------------------
# Create CSV file
# ---------------------------------------------------------

# Output file path for Azure Automation worker
$CsvPath = "$env:TEMP\EntraID_User_SignIns.csv"

# ---------------------------------------------------------
# Fetch users
# ---------------------------------------------------------

$Users = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, SignInActivity, AccountEnabled, CreatedDateTime

Write-Output "Users fetched from Entra ID: $($Users.Count)"

# Remove excluded users based on the HashSet of excluded IDs
$Users = $Users | Where-Object { -not $ExcludedUserIds.Contains([string]$_.Id) }

Write-Output "Users after exclusions applied: $($Users.Count)"

# ---------------------------------------------------------
# Filter logic for inactivity
# ---------------------------------------------------------
# Users are included only if:
# - Their latest sign-in is older than 90 days, OR
# - They have never signed in (no sign-in timestamps)

$ThresholdDate = (Get-Date).AddDays(-90)

$Report = foreach ($User in $Users) {

    # SignInActivity can include both interactive and non-interactive sign-ins.
    $InteractiveDate = $User.SignInActivity.LastSignInDateTime
    $NonInteractiveDate = $User.SignInActivity.LastNonInteractiveSignInDateTime

    # Determine the true latest sign-in date
    $LatestSignIn = $null

    if ($InteractiveDate -and $NonInteractiveDate) {
        $LatestSignIn = if ($InteractiveDate -gt $NonInteractiveDate) { $InteractiveDate } else { $NonInteractiveDate }
    } elseif ($InteractiveDate) {
        $LatestSignIn = $InteractiveDate
    } elseif ($NonInteractiveDate) {
        $LatestSignIn = $NonInteractiveDate
    }

    # Exclude users who HAVE signed in more recently than the threshold date
    # Keep users where LatestSignIn is $null (meaning they have NEVER signed in)
    if ($LatestSignIn -and $LatestSignIn -gt $ThresholdDate) {
        continue
    }

    [PSCustomObject]@{
        DisplayName              = $User.DisplayName
        Email                    = $User.UserPrincipalName
        AccountStatus            = if ($User.AccountEnabled -eq $true) { "Enabled" } else { "Disabled" }
        CreationDate             = if ($User.CreatedDateTime) { $User.CreatedDateTime.ToString("dd-MM-yyyy") } else { "Unknown" }
        LastInteractiveSignIn    = if ($InteractiveDate) { $InteractiveDate.ToString("dd-MM-yyyy") } else { "Never" }
        LastNonInteractiveSignIn = if ($NonInteractiveDate) { $NonInteractiveDate.ToString("dd-MM-yyyy") } else { "Never" }
    }
}

# Export the report to CSV (semicolon delimiter typical in DK/Excel locales)
$Report | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8BOM -Delimiter ';'

# ---------------------------------------------------------
# Send Email
# ---------------------------------------------------------

$SendFrom = "systemalerts@contoso.com"
$EmailTo = "your@email.com"
$Subject = "Monthly Entra ID User Report"
$BodyContent = "Here is the user sign-in report generated on $(Get-Date)."

# Convert CSV to Base64 and attach (Graph fileAttachment requires Base64)
if (Test-Path $CsvPath) {
    $FileBytes  = [System.IO.File]::ReadAllBytes($CsvPath)
    $FileBase64 = [System.Convert]::ToBase64String($FileBytes)
    $FileName   = [System.IO.Path]::GetFileName($CsvPath)

    $AttachmentList = @(
        @{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            Name          = $FileName
            ContentType   = "text/csv"
            ContentBytes  = $FileBase64
        }
    )
} else {
    Write-Warning "File not found at $CsvPath. Sending email without attachment."
    $AttachmentList = @()
}

$Message = @{
    Subject = $Subject
    Body = @{
        ContentType = "HTML"
        Content     = $BodyContent
    }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = $EmailTo } }
    )
    Attachments = $AttachmentList
}

# Send as the mailbox specified in $SendFrom
Send-MgUserMail -UserId $SendFrom -Message $Message -SaveToSentItems -ErrorAction Stop