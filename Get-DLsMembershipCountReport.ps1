# =============================================================================
# Get-GroupMemberReport.ps1
# Distribution List / M365 Groups Member Count Report
# 
# Usage:
#   .\Get-GroupMemberReport.ps1 -DL                  # Only Distribution Lists
#   .\Get-GroupMemberReport.ps1 -M365                # Only Microsoft 365 Groups
#   .\Get-GroupMemberReport.ps1 -All                 # Both DLs and M365 Groups
#   .\Get-GroupMemberReport.ps1 -DL -TestLimit 100   # Test with first ~100 DLs
# =============================================================================

[CmdletBinding(DefaultParameterSetName = 'DL')]
param(
    [Parameter(ParameterSetName = 'DL')]
    [switch]$DL,
    
    [Parameter(ParameterSetName = 'M365')]
    [switch]$M365,
    
    [Parameter(ParameterSetName = 'All')]
    [switch]$All,
    
    [Parameter()]
    [int]$TestLimit = 0  # 0 = no limit, process all
)

# -----------------------------------------------------------------------------
# Determine Mode
# -----------------------------------------------------------------------------

# If no switch specified, default to DL
if (-not $DL -and -not $M365 -and -not $All) {
    $DL = $true
}

$mode = if ($All) { "All" } elseif ($M365) { "M365" } else { "DL" }
Write-Host "`n=== Group Member Report ===" -ForegroundColor Cyan
Write-Host "Mode: $mode $(if ($TestLimit -gt 0) { "(TestLimit: $TestLimit)" })" -ForegroundColor Cyan

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
    Connect-MgGraph -Scopes "Group.Read.All", "Directory.Read.All"
}

# -----------------------------------------------------------------------------
# Fetch Groups via Graph API
# -----------------------------------------------------------------------------

Write-Host "`nFetching groups via Graph API..." -ForegroundColor Cyan

# Optimize: if TestLimit is set, don't fetch all groups
# We fetch 3x TestLimit as buffer (because we filter by type locally)
if ($TestLimit -gt 0) {
    $fetchLimit = if ($All) { $TestLimit } else { $TestLimit * 3 }
    
    Write-Host "  Fetching up to $fetchLimit groups (optimized for TestLimit)..." -ForegroundColor Yellow
    
    $allGroups = Get-MgGroup -Filter "mailEnabled eq true" `
                             -Top $fetchLimit `
                             -Property Id, DisplayName, Mail, GroupTypes `
                             -ConsistencyLevel eventual
}
else {
    Write-Host "  Fetching all groups..." -ForegroundColor Gray
    
    $allGroups = Get-MgGroup -Filter "mailEnabled eq true" `
                             -All `
                             -Property Id, DisplayName, Mail, GroupTypes `
                             -ConsistencyLevel eventual
}

# Separate DLs from M365 Groups based on GroupTypes property
# DL = empty GroupTypes array
# M365 = GroupTypes contains "Unified"
$dlGroups = $allGroups | Where-Object { $_.GroupTypes.Count -eq 0 }
$m365Groups = $allGroups | Where-Object { $_.GroupTypes -contains 'Unified' }

Write-Host "  Distribution Lists fetched:    $($dlGroups.Count)" -ForegroundColor Gray
Write-Host "  Microsoft 365 Groups fetched:  $($m365Groups.Count)" -ForegroundColor Gray

# Select groups based on chosen mode
$selectedGroups = switch ($mode) {
    "DL"   { $dlGroups }
    "M365" { $m365Groups }
    "All"  { $allGroups }
}

# Apply TestLimit after type filtering
if ($TestLimit -gt 0) {
    $selectedGroups = $selectedGroups | Select-Object -First $TestLimit
}

# Transform to working format
$groupsToProcess = $selectedGroups | Select-Object @{N='PrimarySmtpAddress';E={$_.Mail}}, 
                                                    @{N='GroupId';E={$_.Id}}, 
                                                    DisplayName,
                                                    @{N='GroupType';E={ if ($_.GroupTypes -contains 'Unified') { 'M365' } else { 'DL' } }}

$totalCount = $groupsToProcess.Count
Write-Host "  Groups to process:             $totalCount" -ForegroundColor Green

# Safety check
if ($totalCount -eq 0) {
    Write-Host "`nNo groups found to process. Exiting." -ForegroundColor Yellow
    exit
}

# -----------------------------------------------------------------------------
# Process Member Counts
# -----------------------------------------------------------------------------

Write-Host "`nProcessing member counts..." -ForegroundColor Cyan

$results = [System.Collections.ArrayList]::new()
$currentIndex = 0
$errorCount = 0

foreach ($group in $groupsToProcess) {
    $currentIndex++
    
    Write-Progress -Activity "Getting member counts" `
                   -Status "[$currentIndex / $totalCount] $($group.DisplayName)" `
                   -PercentComplete (($currentIndex / $totalCount) * 100)
    
    # Skip if no Group ID
    if ([string]::IsNullOrWhiteSpace($group.GroupId)) {
        Write-Warning "No ID for '$($group.DisplayName)' - skipping"
        $errorCount++
        continue
    }
    
    try {
        # Get transitive member count (includes nested group members)
        $memberCount = Get-MgGroupTransitiveMemberCount -GroupId $group.GroupId -ConsistencyLevel eventual
        
        # Build result with size bucket flags
        $result = [PSCustomObject]@{
            PrimarySmtpAddress = $group.PrimarySmtpAddress
            DisplayName        = $group.DisplayName
            GroupType          = $group.GroupType
            Members            = $memberCount
            "Empty"            = if ($memberCount -eq 0) { "yes" } else { "no" }
            "Size_1-10"        = if ($memberCount -ge 1 -and $memberCount -le 10) { "yes" } else { "no" }
            "11-100"           = if ($memberCount -ge 11 -and $memberCount -le 100) { "yes" } else { "no" }
            "101-200"          = if ($memberCount -ge 101 -and $memberCount -le 200) { "yes" } else { "no" }
            "201-500"          = if ($memberCount -ge 201 -and $memberCount -le 500) { "yes" } else { "no" }
            "501-1000"         = if ($memberCount -ge 501 -and $memberCount -le 1000) { "yes" } else { "no" }
            "1001-5000"        = if ($memberCount -ge 1001 -and $memberCount -le 5000) { "yes" } else { "no" }
            "5000+"            = if ($memberCount -gt 5000) { "yes" } else { "no" }
        }
        
        [void]$results.Add($result)
    }
    catch {
        Write-Warning "Error for '$($group.DisplayName)': $($_.Exception.Message)"
        $errorCount++
        
        # Add error entry for tracking
        $result = [PSCustomObject]@{
            PrimarySmtpAddress = $group.PrimarySmtpAddress
            DisplayName        = $group.DisplayName
            GroupType          = $group.GroupType
            Members            = "ERROR"
            "Empty"            = "error"
            "1-10"             = "error"
            "11-100"           = "error"
            "101-200"          = "error"
            "201-500"          = "error"
            "501-1000"         = "error"
            "1001-5000"        = "error"
            "5000+"            = "error"
        }
        
        [void]$results.Add($result)
    }
}

Write-Progress -Activity "Getting member counts" -Completed

# -----------------------------------------------------------------------------
# Export Results
# -----------------------------------------------------------------------------

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputPath = "GroupMemberReport_${mode}_$timestamp.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  REPORT COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode:             $mode"
Write-Host "  Groups processed: $($results.Count - $errorCount)"
Write-Host "  Errors:           $errorCount"
Write-Host "  Output:           $outputPath"

# -----------------------------------------------------------------------------
# Summary Statistics
# -----------------------------------------------------------------------------

$validResults = $results | Where-Object { $_.Members -ne "ERROR" }

# Show breakdown by group type in All mode
if ($mode -eq "All") {
    Write-Host "`n--- By Group Type ---" -ForegroundColor Yellow
    Write-Host "  Distribution Lists:    $(($validResults | Where-Object { $_.GroupType -eq 'DL' }).Count)"
    Write-Host "  Microsoft 365 Groups:  $(($validResults | Where-Object { $_.GroupType -eq 'M365' }).Count)"
}

Write-Host "`n--- Size Distribution ---" -ForegroundColor Yellow

$stats = @(
    @{ Name = "Empty (0)";   Count = ($validResults | Where-Object { $_.'Empty' -eq 'yes' }).Count }
    @{ Name = "1-10";        Count = ($validResults | Where-Object { $_.'1-10' -eq 'yes' }).Count }
    @{ Name = "11-100";      Count = ($validResults | Where-Object { $_.'11-100' -eq 'yes' }).Count }
    @{ Name = "101-200";     Count = ($validResults | Where-Object { $_.'101-200' -eq 'yes' }).Count }
    @{ Name = "201-500";     Count = ($validResults | Where-Object { $_.'201-500' -eq 'yes' }).Count }
    @{ Name = "501-1000";    Count = ($validResults | Where-Object { $_.'501-1000' -eq 'yes' }).Count }
    @{ Name = "1001-5000";   Count = ($validResults | Where-Object { $_.'1001-5000' -eq 'yes' }).Count }
    @{ Name = "5000+";       Count = ($validResults | Where-Object { $_.'5000+' -eq 'yes' }).Count }
)

foreach ($stat in $stats) {
    Write-Host ("  {0,-15} {1,6}" -f $stat.Name, $stat.Count)
}

# Top 5 largest groups
Write-Host "`n--- Top 5 Largest ---" -ForegroundColor Yellow
$validResults | 
    Sort-Object { [int]$_.Members } -Descending | 
    Select-Object -First 5 | 
    ForEach-Object {
        Write-Host ("  {0,6} members - [{1}] {2}" -f $_.Members, $_.GroupType, $_.PrimarySmtpAddress)
    }

Write-Host ""
