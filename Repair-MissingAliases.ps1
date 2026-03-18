<#
.SYNOPSIS
    Adds missing old-domain aliases to users whose proxyAddresses were not fully
    updated during the domain migration.

.DESCRIPTION
    Reads the pre-migration backup CSV produced by Switch-EmailDomain.ps1 to
    determine each user's old UPN. For each user, checks whether a proxy address
    alias matching their old UPN exists. If missing, adds it as a secondary (smtp:)
    alias so mail sent to the old address is still delivered.

.PARAMETER BackupCsvPath
    Path to the PreMigrationBackup CSV file from the original migration run.

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\Repair-MissingAliases.ps1 -BackupCsvPath "C:\PreMigrationBackup_20260318_153135.csv"

.EXAMPLE
    .\Repair-MissingAliases.ps1 -BackupCsvPath "C:\PreMigrationBackup_20260318_153135.csv" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to the pre-migration backup CSV")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$BackupCsvPath
)

#Requires -Modules ActiveDirectory

# ============================================================================
#  LOGGING
# ============================================================================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile   = Join-Path $scriptDir "AliasRemediation_$timestamp.log"

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARN"    { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry -ForegroundColor Cyan }
    }
    Add-Content -Path $logFile -Value $entry
}

# ============================================================================
#  INITIALISATION
# ============================================================================
Write-Log "Alias Remediation Script Started"
Write-Log "Backup CSV : $BackupCsvPath"
Write-Log "WhatIf     : $WhatIfPreference"
Write-Log "Log File   : $logFile"

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory module loaded"
} catch {
    Write-Log "Failed to load ActiveDirectory module: $_" -Level ERROR
    exit 1
}

# ============================================================================
#  LOAD BACKUP AND ASSESS
# ============================================================================
try {
    $backupUsers = @(Import-Csv -Path $BackupCsvPath -ErrorAction Stop)
    Write-Log "Loaded $($backupUsers.Count) user(s) from backup CSV"
} catch {
    Write-Log "Failed to read backup CSV: $_" -Level ERROR
    exit 1
}

if ($backupUsers.Count -eq 0) {
    Write-Log "Backup CSV is empty. Nothing to do." -Level WARN
    exit 0
}

# ============================================================================
#  SCAN FOR MISSING ALIASES
# ============================================================================
Write-Log "Scanning for missing old-domain aliases..."

$usersToFix = @()

foreach ($row in $backupUsers) {
    $sam    = $row.SamAccountName
    $oldUPN = $row.UserPrincipalName

    if (-not $oldUPN -or $oldUPN -notmatch '@') {
        Write-Log "  $sam - no valid UPN in backup, skipping" -Level WARN
        continue
    }

    $upnPrefix = ($oldUPN -split "@")[0]
    $oldDomain = ($oldUPN -split "@")[1]
    $expectedAlias = "smtp:$upnPrefix@$oldDomain"

    # Get current state from AD
    try {
        $adUser = Get-ADUser -Identity $sam -Properties proxyAddresses -ErrorAction Stop
    } catch {
        Write-Log "  $sam - not found in AD, skipping: $_" -Level WARN
        continue
    }

    $currentProxies = @($adUser.proxyAddresses | Where-Object { $_ })

    # Check if alias already exists (case-insensitive)
    $hasAlias = $currentProxies | Where-Object { $_ -ieq $expectedAlias }

    # Also check if it exists as primary (SMTP: uppercase) -- still counts
    if (-not $hasAlias) {
        $hasAlias = $currentProxies | Where-Object { $_ -ieq "SMTP:$upnPrefix@$oldDomain" }
    }

    if (-not $hasAlias) {
        $usersToFix += [PSCustomObject]@{
            SamAccountName  = $sam
            DisplayName     = $row.DisplayName
            OldUPN          = $oldUPN
            MissingAlias    = $expectedAlias
            CurrentProxies  = $currentProxies
            DN              = $adUser.DistinguishedName
        }
    }
}

Write-Log "Scan complete: $($usersToFix.Count) user(s) missing old-domain alias"

if ($usersToFix.Count -eq 0) {
    Write-Log "All users already have their old-domain alias. Nothing to remediate." -Level SUCCESS
    exit 0
}

# ============================================================================
#  DISPLAY AND CONFIRM
# ============================================================================
Write-Host ""
Write-Log "Users requiring remediation:"
Write-Host ""
Write-Host ("{0,-5} {1,-30} {2,-40} {3}" -f "#", "DisplayName", "Old UPN", "Alias to Add") -ForegroundColor White
Write-Host ("{0,-5} {1,-30} {2,-40} {3}" -f "-"*5, "-"*30, "-"*40, "-"*40) -ForegroundColor Gray

$i = 0
foreach ($u in $usersToFix) {
    $i++
    Write-Host ("{0,-5} {1,-30} {2,-40} {3}" -f $i, $u.DisplayName, $u.OldUPN, $u.MissingAlias)
}

Write-Host ""

if (-not $WhatIfPreference) {
    $confirm = Read-Host "Add missing alias for $($usersToFix.Count) user(s)? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Log "Remediation cancelled by user" -Level WARN
        exit 0
    }
}

# ============================================================================
#  APPLY FIXES
# ============================================================================
$successCount = 0
$errorCount   = 0
$total        = $usersToFix.Count
$counter      = 0

foreach ($u in $usersToFix) {
    $counter++
    $label = "$($u.DisplayName) ($($u.SamAccountName))"

    Write-Log "[$counter/$total] $label"
    Write-Log "  Adding alias: $($u.MissingAlias)"

    try {
        if ($PSCmdlet.ShouldProcess($label, "Add alias '$($u.MissingAlias)'")) {
            Set-ADUser -Identity $u.DN -Add @{ proxyAddresses = $u.MissingAlias } -ErrorAction Stop
            Write-Log "  Alias added" -Level SUCCESS
            $successCount++
        }
    } catch {
        Write-Log "  FAILED: $_" -Level ERROR
        $errorCount++
    }
}

# ============================================================================
#  SUMMARY
# ============================================================================
Write-Host ""
Write-Log "===================================================================="
Write-Log "REMEDIATION COMPLETE"
Write-Log "  Total       : $total"
Write-Log "  Successful  : $successCount" -Level SUCCESS
if ($errorCount -gt 0) {
    Write-Log "  Failed      : $errorCount" -Level ERROR
} else {
    Write-Log "  Failed      : 0"
}
Write-Log "  Log File    : $logFile"
Write-Log "===================================================================="
Write-Log "Script finished"
