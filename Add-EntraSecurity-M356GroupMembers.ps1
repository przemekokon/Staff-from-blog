<#
.SYNOPSIS
    Adds users from a TXT file to an Entra ID Group via Microsoft Graph.

.DESCRIPTION
    The script reads UPNs from a text file, resolves each user (first by UPN,
    then by mail as fallback), and adds them to the specified group.
    Works with both Security Groups and Microsoft 365 Groups.

.PARAMETER GroupId
    Object ID of the target Entra ID Group.

.PARAMETER FilePath
    Path to the TXT file containing one UPN per line.

.EXAMPLE
    .\Add-GroupMembers.ps1 -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -FilePath ".\users.txt"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, HelpMessage = "Object ID of the target Entra ID Group")]
    [string]$GroupId,

    [Parameter(Mandatory, HelpMessage = "Path to the TXT file with UPNs (one per line)")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$FilePath
)

# --- Connect to Graph if not already connected ---
try {
    $null = Get-MgContext -ErrorAction Stop
}
catch {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All", "Directory.ReadWrite.All"
}

# --- Validate that the group exists before processing ---
Write-Host "`nValidating group ID..." -ForegroundColor Cyan
try {
    $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    Write-Host "Target group: $($group.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Error "Could not find group with ID: $GroupId. Aborting."
    exit 1
}

# --- Read UPNs from file, skip empty lines ---
$upnList = Get-Content $FilePath | Where-Object { $_.Trim() -ne "" }
Write-Host "Users to process: $($upnList.Count)`n" -ForegroundColor Cyan

# --- Counters for summary ---
$added    = 0
$skipped  = 0
$notFound = 0

foreach ($entry in $upnList) {

    $entry = $entry.Trim()
    $user  = $null

    # Step 1: Try to resolve user by UPN
    $user = Get-MgUser -UserId $entry -ErrorAction SilentlyContinue

    # Step 2: Fallback - try to find user by mail attribute
    if (-not $user) {
        Write-Warning "UPN not found: '$entry' - trying mail fallback..."
        $user = Get-MgUser -Filter "mail eq '$entry'" -ErrorAction SilentlyContinue |
                Select-Object -First 1
    }

    # Step 3: Still not found - log and skip
    if (-not $user) {
        Write-Warning "User not found (UPN or mail): '$entry' - skipping."
        $notFound++
        continue
    }

    # Step 4: Check if user is already a member (avoids errors on duplicate add)
    $isMember = Get-MgGroupMember -GroupId $GroupId -Filter "id eq '$($user.Id)'" -ErrorAction SilentlyContinue
    if ($isMember) {
        Write-Host "  [SKIP] $($user.UserPrincipalName) is already a member." -ForegroundColor Yellow
        $skipped++
        continue
    }

    # Step 5: Add user to the group
    try {
        New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"
        }
        Write-Host "  [OK]   $($user.UserPrincipalName) added successfully." -ForegroundColor Green
        $added++
    }
    catch {
        Write-Warning "Failed to add $($user.UserPrincipalName): $_"
        $notFound++
    }
}

# --- Summary ---
Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "Added:     $added"    -ForegroundColor Green
Write-Host "Skipped:   $skipped"  -ForegroundColor Yellow
Write-Host "Not found: $notFound" -ForegroundColor Red