# =============================================================================
# Get-UserGroupMemberships.ps1
#
# Reads UPNs from a text file, fetches group memberships for each user
# via Microsoft Graph, and exports a flat CSV report.
#
# Check line 40 for exclusions.
#
# Input:
#   - users.txt (one UPN per line, same directory as script)
#
# Output:
#   - UserGroupMemberships_<timestamp>.csv
#
# Columns:
#   UserPrincipalName, PrimarySmtpAddress, GroupName, GroupEmail, GroupType
#
# Usage:
#   .\Get-UserGroupMemberships.ps1
#   .\Get-UserGroupMemberships.ps1 -InputFile "C:\path\to\myusers.txt"
# =============================================================================

param(
    [string]$InputFile
)

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

if (-not $InputFile) {
    $InputFile = Join-Path $scriptPath "users.txt"
}

$reportPath = Join-Path $scriptPath "UserGroupMemberships_$timestamp.csv"

# Groups to exclude from the report (matched by DisplayName, case-insensitive)
$excludedGroups = @(
    'Tenant Guests'
)

# -----------------------------------------------------------------------------
# Helper: Determine group type from Graph properties
# -----------------------------------------------------------------------------

function Get-GroupType {
    param(
        [bool]$SecurityEnabled,
        [bool]$MailEnabled,
        [string[]]$GroupTypes
    )

    # M365 Group (Unified)
    if ($GroupTypes -contains 'Unified') {
        return 'M365Group'
    }

    # Mail-enabled Security Group
    if ($MailEnabled -and $SecurityEnabled) {
        return 'MailSecurityGroup'
    }

    # Distribution List (mail-enabled, not security)
    if ($MailEnabled -and -not $SecurityEnabled) {
        return 'DistributionList'
    }

    # Entra Security Group (security only, no mail)
    if ($SecurityEnabled -and -not $MailEnabled) {
        return 'EntraSecurityGroup'
    }

    return 'Other'
}

# -----------------------------------------------------------------------------
# Connection Check
# -----------------------------------------------------------------------------

Write-Host "`n=== User Group Memberships Report ===" -ForegroundColor Cyan
Write-Host "Checking Microsoft Graph connection..." -ForegroundColor Cyan

try {
    $mgContext = Get-MgContext -ErrorAction Stop
    if (-not $mgContext) { throw "No session" }
    Write-Host "  Connected as $($mgContext.Account)" -ForegroundColor Green
}
catch {
    Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All", "GroupMember.Read.All", "Group.Read.All"
}

# -----------------------------------------------------------------------------
# Load Input File
# -----------------------------------------------------------------------------

if (-not (Test-Path $InputFile)) {
    Write-Host "  ERROR: Input file not found: $InputFile" -ForegroundColor Red
    return
}

$upnList = Get-Content $InputFile | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
Write-Host "  Loaded $($upnList.Count) UPNs from $InputFile" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Process Each User (incremental CSV write every 100 users)
# -----------------------------------------------------------------------------

$totalUsers = $upnList.Count
$currentIndex = 0
$errorCount = 0
$processedCount = 0
$totalRows = 0
$buffer = [System.Collections.ArrayList]::new()
$fileCreated = $false

foreach ($upn in $upnList) {
    $currentIndex++
    Write-Progress -Activity "Processing users" `
                   -Status "[$currentIndex / $totalUsers] $upn" `
                   -PercentComplete (($currentIndex / $totalUsers) * 100)

    # --- Get user object (resolves UPN to ID + mail) ---
    try {
        $user = Get-MgUser -UserId $upn -Property Id, UserPrincipalName, Mail -ErrorAction Stop
    }
    catch {
        Write-Host "  SKIP: Cannot find user $upn - $($_.Exception.Message)" -ForegroundColor Yellow
        $errorCount++
        continue
    }

    $userMail = if ($user.Mail) { $user.Mail } else { "" }
    $processedCount++

    # --- Get all group memberships ---
    try {
        $memberships = Get-MgUserMemberOf -UserId $user.Id -All
    }
    catch {
        Write-Host "  SKIP: Cannot get memberships for $upn - $($_.Exception.Message)" -ForegroundColor Yellow
        $errorCount++
        continue
    }

    # --- Filter to groups only (skip directory roles, admin units etc.) ---
    # Wrap in @() to ensure array even for single result (PS 5.1 safety)
    $groups = @($memberships | Where-Object {
        $_.'@odata.type' -eq '#microsoft.graph.group' -or
        $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'
    })

    if ($groups.Count -eq 0) {
        [void]$buffer.Add([PSCustomObject]@{
            UserPrincipalName  = $user.UserPrincipalName
            PrimarySmtpAddress = $userMail
            GroupName          = ""
            GroupEmail         = ""
            GroupType          = "(no groups)"
        })
    }
    else {
        foreach ($grp in $groups) {
            $props        = $grp.AdditionalProperties
            $displayName  = if ($grp.DisplayName)  { $grp.DisplayName }  else { $props['displayName'] }

            # Skip excluded groups
            if ($excludedGroups -contains $displayName) { continue }

            $mail         = if ($grp.Mail)         { $grp.Mail }         else { $props['mail'] }
            $secEnabled   = if ($null -ne $grp.SecurityEnabled) { $grp.SecurityEnabled } else { $props['securityEnabled'] }
            $mailEnabled  = if ($null -ne $grp.MailEnabled)     { $grp.MailEnabled }     else { $props['mailEnabled'] }
            $groupTypes   = if ($grp.GroupTypes)   { $grp.GroupTypes }   else { $props['groupTypes'] }

            if ($null -eq $groupTypes) { $groupTypes = @() }

            $groupType = Get-GroupType -SecurityEnabled ([bool]$secEnabled) `
                                       -MailEnabled ([bool]$mailEnabled) `
                                       -GroupTypes $groupTypes

            [void]$buffer.Add([PSCustomObject]@{
                UserPrincipalName  = $user.UserPrincipalName
                PrimarySmtpAddress = $userMail
                GroupName          = $displayName
                GroupEmail         = $mail
                GroupType          = $groupType
            })
        }
    }

    # --- Flush buffer every 100 users ---
    if ($currentIndex % 100 -eq 0 -and $buffer.Count -gt 0) {
        if (-not $fileCreated) {
            $buffer | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
            $fileCreated = $true
        }
        else {
            $buffer | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8 -Append
        }
        $totalRows += $buffer.Count
        $buffer.Clear()
    }
}

# --- Flush remaining rows ---
if ($buffer.Count -gt 0) {
    if (-not $fileCreated) {
        $buffer | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        $fileCreated = $true
    }
    else {
        $buffer | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8 -Append
    }
    $totalRows += $buffer.Count
    $buffer.Clear()
}

Write-Progress -Activity "Processing users" -Completed

# -----------------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  REPORT COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Users processed:   $processedCount / $totalUsers"
Write-Host "  Errors/skipped:    $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Total rows in CSV: $totalRows"
Write-Host "`n  Output: $reportPath" -ForegroundColor Cyan
Write-Host ""
