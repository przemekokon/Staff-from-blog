# =============================================================================
# Set-UserLicenseCleanup.ps1
# Removes all licenses from user(s) and optionally assigns a specific license
#
# Usage:
#   .\Set-UserLicenseCleanup.ps1 -User przemek@contoso.com -Remove
#   .\Set-UserLicenseCleanup.ps1 -User przemek@contoso.com -Update
#   .\Set-UserLicenseCleanup.ps1 -Users C:\users.txt -Remove
#   .\Set-UserLicenseCleanup.ps1 -Users C:\users.txt -Update
# =============================================================================

[CmdletBinding()]
param(
    # Single user UPN
    [Parameter(ParameterSetName = 'SingleRemove', Mandatory)]
    [Parameter(ParameterSetName = 'SingleUpdate', Mandatory)]
    [string]$User,

    # Path to TXT file with UPNs (one per line)
    [Parameter(ParameterSetName = 'BulkRemove', Mandatory)]
    [Parameter(ParameterSetName = 'BulkUpdate', Mandatory)]
    [string]$Users,

    # Remove all licenses only
    [Parameter(ParameterSetName = 'SingleRemove', Mandatory)]
    [Parameter(ParameterSetName = 'BulkRemove', Mandatory)]
    [switch]$Remove,

    # Remove all licenses and assign the target license
    [Parameter(ParameterSetName = 'SingleUpdate', Mandatory)]
    [Parameter(ParameterSetName = 'BulkUpdate', Mandatory)]
    [switch]$Update
)

# =============================================================================
# TARGET LICENSE - change this SkuId to the license you want to assign
# Find SkuIds via: Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId
$targetSkuId = "f30db892-...-80727f46fd3d"
# =============================================================================

# -----------------------------------------------------------------------------
# Connection Check
# -----------------------------------------------------------------------------

Write-Host "`nChecking Microsoft Graph connection..." -ForegroundColor Cyan

try {
    $mgContext = Get-MgContext -ErrorAction Stop
    if (-not $mgContext) { throw "No session" }
    Write-Host "  Connected as $($mgContext.Account)" -ForegroundColor Green
}
catch {
    Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Organization.Read.All", "User.ReadWrite.All"
}

# -----------------------------------------------------------------------------
# Load UPNs
# -----------------------------------------------------------------------------

if ($User) {
    $upns = @($User)
}
else {
    $upns = Get-Content -Path $Users | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

$totalCount = $upns.Count
$mode       = if ($Update) { "UPDATE (remove all + assign target license)" } else { "REMOVE (remove all licenses)" }

# -----------------------------------------------------------------------------
# Confirmation Prompt
# -----------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "  Mode:    $mode" -ForegroundColor Yellow
Write-Host "  Users:   $totalCount" -ForegroundColor Yellow
if ($Update) {
    Write-Host "  Target license SkuId: $targetSkuId" -ForegroundColor Yellow
}
Write-Host "========================================" -ForegroundColor Yellow

$confirm = Read-Host "`nAre you sure you want to proceed? (yes/no)"
if ($confirm -ne "yes") {
    Write-Host "Aborted." -ForegroundColor Red
    exit
}

# -----------------------------------------------------------------------------
# Process Users
# -----------------------------------------------------------------------------

$results    = [System.Collections.ArrayList]::new()
$current    = 0
$errorCount = 0

foreach ($upn in $upns) {
    $current++

    Write-Progress -Activity "Processing licenses" `
                   -Status "[$current / $totalCount] $upn" `
                   -PercentComplete (($current / $totalCount) * 100)

    try {
        # Get user and their current licenses
        $mgUser           = Get-MgUser -UserId $upn -Property "Id,UserPrincipalName,AssignedLicenses" -ErrorAction Stop
        $assignedLicenses = $mgUser.AssignedLicenses
        $removedSkuIds    = $assignedLicenses | ForEach-Object { $_.SkuId }

        if ($removedSkuIds.Count -eq 0) {
            Write-Host "  No licenses found for $upn - skipping" -ForegroundColor Gray

            [void]$results.Add([PSCustomObject]@{
                UserPrincipalName = $upn
                Status            = "SKIPPED - no licenses assigned"
                LicensesRemoved   = "none"
                LicenseAssigned   = "n/a"
            })
            continue
        }

        # Remove all licenses and optionally assign target license in a single API call
        if ($Update) {
            Set-MgUserLicense -UserId $mgUser.Id `
                              -RemoveLicenses $removedSkuIds `
                              -AddLicenses @(@{ SkuId = $targetSkuId }) `
                              -ErrorAction Stop | Out-Null
        }
        else {
            Set-MgUserLicense -UserId $mgUser.Id `
                              -RemoveLicenses $removedSkuIds `
                              -AddLicenses @() `
                              -ErrorAction Stop | Out-Null
        }

        [void]$results.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            Status            = "OK"
            LicensesRemoved   = ($removedSkuIds -join ", ")
            LicenseAssigned   = if ($Update) { $targetSkuId } else { "n/a" }
        })
    }
    catch {
        Write-Warning "Error for '$upn': $($_.Exception.Message)"
        $errorCount++

        [void]$results.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            Status            = "ERROR: $($_.Exception.Message)"
            LicensesRemoved   = "n/a"
            LicenseAssigned   = "n/a"
        })
    }
}

Write-Progress -Activity "Processing licenses" -Completed

# -----------------------------------------------------------------------------
# Export Results
# -----------------------------------------------------------------------------

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = ".\LicenseCleanup_$timestamp.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding Default

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Processed: $($results.Count - $errorCount)"
Write-Host "  Errors:    $errorCount"
Write-Host "  Output:    $outputPath"
Write-Host ""
