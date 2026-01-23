<#
.SYNOPSIS
Creates and configures a Microsoft Teams Room resource account.

.DESCRIPTION
Creates an Exchange room mailbox, configures Places metadata,
sets booking policies, assigns Teams Rooms licenses, and prepares
the account for Microsoft Teams Rooms devices.

.REQUIREMENTS
PowerShell modules:
- ExchangeOnlineManagement
- Microsoft.Graph.Users
- Microsoft.Graph.Identity.DirectoryManagement

Required permissions:
- Exchange Administrator
- User Administrator

Tenant prerequisites:
- Microsoft Teams Rooms license available (Pro, Pro without Audio Conferencing, or Standard)
- Exchange Online configured
- Microsoft Graph access enabled

.AUTHOR
Peter Busted

.NOTES
Test in non-production before use.
Exclude Teams Room resource accounts from interactive MFA policies.
#>

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURATION VARIABLES
# ============================================================================

# Display name and primary SMTP address for the room
$RoomName    = "Teams Room - Test4"
$PrimarySmtp = "teamsroom-test4@contoso.com"
$Alias       = ($PrimarySmtp.Split("@")[0])

# Room Finder metadata
$Capacity    = 12
$Floor       = 1
$City        = "Copenhagen"
$Street      = "Abbey Road 123"
$PostalCode  = "2100"
$Country     = "Denmark"
$Building    = "Copenhagen Teams Rooms"

# Entra ID / licensing configuration
$UsageLocation = "DK"
$Password      = "YourSecurePassword123!"

# Teams Rooms license model
$LicenseMode = "Pro"
# Pro
# ProNoAudio
# Standard

# ============================================================================
# CONNECT TO REQUIRED SERVICES
# ============================================================================

# Connect to Exchange Online (for mailbox and room configuration)
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

# Connect to Microsoft Graph (for Entra ID and licensing operations)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.Read.All" | Out-Null

# ============================================================================
# CREATE ROOM MAILBOX (IF IT DOES NOT ALREADY EXIST)
# ============================================================================

Write-Host "Creating room mailbox" -ForegroundColor Cyan
$SecurePass = ConvertTo-SecureString $Password -AsPlainText -Force

$existing = Get-Mailbox -Identity $PrimarySmtp -ErrorAction SilentlyContinue
if (-not $existing) {
    New-Mailbox -Name $RoomName `
        -Alias $Alias `
        -PrimarySmtpAddress $PrimarySmtp `
        -Room `
        -EnableRoomMailboxAccount $true `
        -RoomMailboxPassword $SecurePass
} else {
    Write-Host "Room mailbox already exists, continuing" -ForegroundColor Yellow
}

# Allow time for mailbox to propagate before applying metadata
Start-Sleep -Seconds 20

# ============================================================================
# CONFIGURE PLACE METADATA FOR ROOM FINDER (CITY, BUILDING, FLOOR, CAPACITY)
# ============================================================================

Write-Host "Setting place metadata" -ForegroundColor Cyan
for ($i = 1; $i -le 6; $i++) {
    try {
        Set-Place -Identity $PrimarySmtp `
            -City $City `
            -Building $Building `
            -Street $Street `
            -PostalCode $PostalCode `
            -CountryOrRegion $Country `
            -Floor $Floor `
            -Capacity $Capacity `
            -WarningAction SilentlyContinue

        # Validate that values are visible in Places backend
        $place = Get-Place -Identity $PrimarySmtp -ErrorAction Stop
        if ($place.City -eq $City -and $place.Building -eq $Building) {
            break
        }

        throw "Place values not visible yet"
    }
    catch {
        if ($i -eq 6) { throw }
        Start-Sleep -Seconds 10
    }
}

# ============================================================================
# CONFIGURE BOOKING AND AUTO-ACCEPT BEHAVIOR FOR THE ROOM
# ============================================================================

Write-Host "Setting calendar processing rules" -ForegroundColor Cyan
Set-CalendarProcessing -Identity $PrimarySmtp `
  -AutomateProcessing AutoAccept `
  -AllowConflicts $false `
  -BookingWindowInDays 365 `
  -AllBookInPolicy $true `
  -ProcessExternalMeetingMessages $true `
  -AddOrganizerToSubject $false `
  -DeleteSubject $false `
  -DeleteComments $false `
  -RemovePrivateProperty $false `
  -AddAdditionalResponse $false `
  -WarningAction SilentlyContinue

# ============================================================================
# SET REGIONAL SETTINGS (LANGUAGE, TIMEZONE, FOLDER LOCALIZATION)
# ============================================================================

Write-Host "Setting mailbox regional configuration" -ForegroundColor Cyan
Set-MailboxRegionalConfiguration -Identity $PrimarySmtp `
    -Language "da-DK" `
    -TimeZone "Romance Standard Time" `
    -LocalizeDefaultFolderName:$true

# ============================================================================
# RESOLVE ENTRA ID OBJECT ID FROM EXCHANGE MAILBOX
# This ensures Graph operations target the correct Entra object
# ============================================================================

Write-Host "Resolving Entra object id from Exchange" -ForegroundColor Cyan
$mbx = Get-Mailbox -Identity $PrimarySmtp
$id  = $mbx.ExternalDirectoryObjectId
if (-not $id) { throw "ExternalDirectoryObjectId is empty. The mailbox is not linked to an Entra object yet." }

# ============================================================================
# WAIT UNTIL MICROSOFT GRAPH CAN SEE THE ENTRA OBJECT
# ============================================================================

Write-Host "Waiting for Graph to see the Entra object" -ForegroundColor Cyan
$mgUser = $null
for ($i = 1; $i -le 30; $i++) {
    try {
        $mgUser = Get-MgUser -UserId $id -ErrorAction Stop
        break
    } catch {
        Start-Sleep -Seconds 10
    }
}
if (-not $mgUser) { throw "Graph still cannot see the user after waiting. Verify tenant context and replication." }

# ============================================================================
# CONFIGURE ENTRA USER PROPERTIES REQUIRED FOR LICENSING
# ============================================================================

Write-Host "Setting usage location and password expiration policy in Entra" -ForegroundColor Cyan
Update-MgUser -UserId $id -UsageLocation $UsageLocation -PasswordPolicies "DisablePasswordExpiration"

# ============================================================================
# RESOLVE TEAMS ROOMS LICENSE SKU DYNAMICALLY
# ============================================================================

Write-Host "Resolving license sku id dynamically" -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All

switch ($LicenseMode) {
    "Pro"        { $skuPartNumber = "Microsoft_Teams_Rooms_Pro" }
    "ProNoAudio" { $skuPartNumber = "Microsoft_Teams_Rooms_Pro_without_Audio_Conferencing" }
    "Standard"   { $skuPartNumber = "MEETING_ROOM" }
    default      { throw "Invalid LicenseMode value" }
}

$sku = $skus | Where-Object { $_.SkuPartNumber -eq $skuPartNumber } | Select-Object -First 1
if (-not $sku) { throw "SKU not found in tenant: $skuPartNumber" }

# ============================================================================
# ASSIGN TEAMS ROOMS LICENSE TO THE ROOM ACCOUNT
# ============================================================================

Write-Host "Assigning license" -ForegroundColor Cyan
Set-MgUserLicense -UserId $id -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @()

# ============================================================================
# FINAL VERIFICATION OUTPUT
# ============================================================================

Write-Host "Verification" -ForegroundColor Green
Get-Place -Identity $PrimarySmtp | Select *
Get-CalendarProcessing -Identity $PrimarySmtp | Format-List AutomateProcessing,AllowConflicts,BookingWindowInDays,AllBookInPolicy,ProcessExternalMeetingMessages
Get-Mailbox -Identity $PrimarySmtp | Format-List Name,PrimarySmtpAddress,ExternalDirectoryObjectId
Get-MgUser -UserId $id | Format-List Id,UserPrincipalName,Mail,UsageLocation
Get-MgUserLicenseDetail -UserId $id | Select-Object SkuPartNumber
