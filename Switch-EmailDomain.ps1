<#
.SYNOPSIS
    Updates UPN and primary email address for AD users during a Microsoft 365 domain migration.

.DESCRIPTION
    For each user in the specified OU this script will:
    - Update the User Principal Name (UPN) to the new domain
    - Update proxyAddresses: set new domain as primary SMTP, demote old primary to alias
    - Update the mail attribute to match the new primary SMTP address

    The old domain is auto-detected from existing user UPNs. A pre-change backup CSV
    is always exported before any modifications are made.

.PARAMETER NewDomain
    The new email domain (e.g., "contoso.com"). Do not include the @ symbol.

.PARAMETER OUName
    The name (or partial name) of the OU to target. If multiple OUs match, you will
    be prompted to choose one.

.PARAMETER IncludeDisabledUsers
    By default disabled accounts are skipped. Use this switch to include them.

.PARAMETER WhatIf
    Preview all changes without applying them.

.EXAMPLE
    .\Switch-EmailDomain.ps1 -NewDomain "contoso.com" -OUName "Corporate Users"

.EXAMPLE
    .\Switch-EmailDomain.ps1 -NewDomain "contoso.com" -OUName "Sales" -IncludeDisabledUsers -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "New email domain (e.g., contoso.com)")]
    [ValidateNotNullOrEmpty()]
    [string]$NewDomain,

    [Parameter(Mandatory = $true, HelpMessage = "OU name to search for")]
    [ValidateNotNullOrEmpty()]
    [string]$OUName,

    [switch]$IncludeDisabledUsers
)

#Requires -Modules ActiveDirectory

# ===========================================================================
#  LOGGING
# ===========================================================================
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile    = Join-Path $scriptDir "DomainMigration_$timestamp.log"
$backupFile = Join-Path $scriptDir "PreMigrationBackup_$timestamp.csv"

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
    # Bypass WhatIf propagation so the log file is always written
    $entry | Out-File -FilePath $logFile -Append -Encoding utf8 -WhatIf:$false -Confirm:$false
}

# ===========================================================================
#  INITIALISATION
# ===========================================================================
$NewDomain = $NewDomain.TrimStart("@").Trim()

Write-Log "Email Domain Migration Script Started"
Write-Log "New Domain      : $NewDomain"
Write-Log "OU Search       : $OUName"
Write-Log "WhatIf Mode     : $WhatIfPreference"
Write-Log "Include Disabled: $($IncludeDisabledUsers.IsPresent)"
Write-Log "Log File        : $logFile"

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory module loaded"
} catch {
    Write-Log "Failed to load ActiveDirectory module: $_" -Level ERROR
    exit 1
}

# ===========================================================================
#  OU DISCOVERY
# ===========================================================================
Write-Log "Searching for OUs matching '$OUName'..."

try {
    $escapedOUName = $OUName -replace "'", "''"
    $matchingOUs = @(
        Get-ADOrganizationalUnit -Filter "Name -like '*$escapedOUName*'" -Properties Name |
            Select-Object Name, DistinguishedName
    )
} catch {
    Write-Log "OU search failed: $_" -Level ERROR
    exit 1
}

if ($matchingOUs.Count -eq 0) {
    Write-Log "No OUs found matching '$OUName'" -Level ERROR
    exit 1
}

if ($matchingOUs.Count -eq 1) {
    $selectedOU = $matchingOUs[0]
    Write-Log "Found OU: $($selectedOU.Name) ($($selectedOU.DistinguishedName))"
} else {
    Write-Log "Multiple OUs found matching '$OUName'. Please select one:" -Level WARN
    for ($i = 0; $i -lt $matchingOUs.Count; $i++) {
        Write-Host "  [$($i + 1)] $($matchingOUs[$i].Name) - $($matchingOUs[$i].DistinguishedName)" -ForegroundColor White
    }
    do {
        $sel = Read-Host "`nEnter selection (1-$($matchingOUs.Count))"
        $idx = $sel -as [int]
    } while ($idx -lt 1 -or $idx -gt $matchingOUs.Count)

    $selectedOU = $matchingOUs[$idx - 1]
    Write-Log "Selected OU: $($selectedOU.Name) ($($selectedOU.DistinguishedName))"
}

# ===========================================================================
#  RETRIEVE USERS
# ===========================================================================
Write-Log "Retrieving users from: $($selectedOU.DistinguishedName)"

$adProps = @(
    "UserPrincipalName", "mail", "proxyAddresses",
    "SamAccountName", "DisplayName", "Enabled", "DistinguishedName"
)

try {
    $allUsers = @(Get-ADUser -SearchBase $selectedOU.DistinguishedName -Filter * -Properties $adProps)
} catch {
    Write-Log "Failed to retrieve users: $_" -Level ERROR
    exit 1
}

if ($IncludeDisabledUsers) {
    $users = $allUsers
    Write-Log "Total users (including disabled): $($users.Count)"
} else {
    $users         = @($allUsers | Where-Object { $_.Enabled -eq $true })
    $disabledCount = $allUsers.Count - $users.Count
    Write-Log "Enabled users: $($users.Count)  |  Skipped disabled: $disabledCount"
}

if ($users.Count -eq 0) {
    Write-Log "No users found in the selected OU" -Level WARN
    exit 0
}

# Filter out users already on the new domain
$usersToMigrate = @($users | Where-Object {
    ($_.UserPrincipalName -split "@")[1] -ne $NewDomain
})

$skippedAlready = $users.Count - $usersToMigrate.Count
if ($skippedAlready -gt 0) {
    Write-Log "$skippedAlready user(s) already on '$NewDomain' - skipping" -Level WARN
}

if ($usersToMigrate.Count -eq 0) {
    Write-Log "All users are already on the new domain. Nothing to do." -Level WARN
    exit 0
}

# Detect old domain(s)
$oldDomains = $usersToMigrate |
    ForEach-Object { ($_.UserPrincipalName -split "@")[1] } |
    Sort-Object -Unique

Write-Log "Detected old UPN domain(s): $($oldDomains -join ', ')"

# ===========================================================================
#  PRE-CHANGE BACKUP
# ===========================================================================
Write-Log "Exporting pre-migration backup to: $backupFile"

try {
    $usersToMigrate |
        Select-Object SamAccountName, DisplayName, UserPrincipalName, mail, Enabled,
                      @{N = "proxyAddresses"; E = { ($_.proxyAddresses -join ";") }},
                      DistinguishedName |
        Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8

    Write-Log "Backup saved ($($usersToMigrate.Count) users)" -Level SUCCESS
} catch {
    Write-Log "Backup export failed: $_" -Level ERROR
    exit 1
}

# ===========================================================================
#  CONFIRMATION
# ===========================================================================
$modeTag = if ($WhatIfPreference) { "[WHATIF] " } else { "" }

Write-Host ""
Write-Host "${modeTag}====================================================================" -ForegroundColor Magenta
Write-Host "${modeTag}  MIGRATION SUMMARY"                                                  -ForegroundColor Magenta
Write-Host "${modeTag}====================================================================" -ForegroundColor Magenta
Write-Host "${modeTag}  Target OU       : $($selectedOU.DistinguishedName)"                 -ForegroundColor White
Write-Host "${modeTag}  New Domain      : $NewDomain"                                       -ForegroundColor White
Write-Host "${modeTag}  Old Domain(s)   : $($oldDomains -join ', ')"                        -ForegroundColor White
Write-Host "${modeTag}  Users to Update : $($usersToMigrate.Count)"                         -ForegroundColor White
Write-Host "${modeTag}  Disabled Incl.  : $($IncludeDisabledUsers.IsPresent)"               -ForegroundColor White
Write-Host "${modeTag}====================================================================" -ForegroundColor Magenta
Write-Host ""

if (-not $WhatIfPreference) {
    $confirm = Read-Host "Proceed with updating $($usersToMigrate.Count) user(s)? (Y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Log "Migration cancelled by user" -Level WARN
        exit 0
    }
}

# ===========================================================================
#  MIGRATION
# ===========================================================================
$successCount = 0
$errorCount   = 0
$total        = $usersToMigrate.Count
$counter      = 0

foreach ($user in $usersToMigrate) {
    $counter++
    $label = "$($user.DisplayName) ($($user.SamAccountName))"

    Write-Log "----------------------------------------------------------------"
    Write-Log "[$counter/$total] Processing: $label"

    try {
        # -- Derive values --------------------------------------------------
        $currentUPN = $user.UserPrincipalName
        $upnPrefix  = ($currentUPN -split "@")[0]
        $oldDomain  = ($currentUPN -split "@")[1]
        $newUPN     = "$upnPrefix@$NewDomain"
        $newEmail   = "$upnPrefix@$NewDomain"

        # -- 1. Update UPN --------------------------------------------------
        Write-Log "  UPN: $currentUPN -> $newUPN"

        if ($PSCmdlet.ShouldProcess($label, "Set UPN to '$newUPN'")) {
            Set-ADUser -Identity $user.DistinguishedName -UserPrincipalName $newUPN -ErrorAction Stop
            Write-Log "  UPN updated" -Level SUCCESS
        }

        # -- 2. Update proxyAddresses ---------------------------------------
        $currentProxies = @($user.proxyAddresses | Where-Object { $_ })

        if ($currentProxies.Count -eq 0) {
            # No proxy addresses exist - build from scratch
            Write-Log "  proxyAddresses: (empty) - creating entries"

            $newProxies = @(
                "SMTP:$newEmail"
                "smtp:$upnPrefix@$oldDomain"
            )

            Write-Log "  + SMTP:$newEmail (primary)"
            Write-Log "  + smtp:$upnPrefix@$oldDomain (alias)"
        } else {
            Write-Log "  proxyAddresses: $($currentProxies.Count) existing entries"

            $newProxies = [System.Collections.ArrayList]@()

            # Look for an existing entry at the new domain we can promote
            $newDomainEntries = @($currentProxies | Where-Object {
                $_ -match "(?i)@$([regex]::Escape($NewDomain))$"
            })

            # Prefer exact match on UPN prefix; fall back to first new-domain entry
            $entryToPromote = $newDomainEntries |
                Where-Object { ($_ -split ":", 2)[1] -ieq $newEmail } |
                Select-Object -First 1

            if (-not $entryToPromote -and $newDomainEntries.Count -gt 0) {
                $entryToPromote = $newDomainEntries[0]
            }

            foreach ($proxy in $currentProxies) {
                if ($entryToPromote -and $proxy -ceq $entryToPromote) {
                    # Promote existing new-domain alias to primary
                    $addr = ($proxy -split ":", 2)[1]
                    [void]$newProxies.Add("SMTP:$addr")
                    Write-Log "  ^ Promoted existing alias to primary: SMTP:$addr"
                }
                elseif ($proxy -cmatch '^SMTP:') {
                    # Demote current primary to alias
                    $addr = ($proxy -split ":", 2)[1]
                    [void]$newProxies.Add("smtp:$addr")
                    Write-Log "  v Demoted old primary to alias: smtp:$addr"
                }
                else {
                    # Keep everything else as-is (aliases, X500, SIP, etc.)
                    [void]$newProxies.Add($proxy)
                }
            }

            # If no existing new-domain entry was found, add a new primary
            if (-not $entryToPromote) {
                [void]$newProxies.Add("SMTP:$newEmail")
                Write-Log "  + Added new primary: SMTP:$newEmail"
            }

            # Ensure an alias exists for the old UPN-based address
            $oldAlias = "smtp:$upnPrefix@$oldDomain"
            $hasOldAlias = $newProxies | Where-Object { $_ -ieq $oldAlias }
            if (-not $hasOldAlias) {
                [void]$newProxies.Add($oldAlias)
                Write-Log "  + Added old UPN alias: $oldAlias"
            }
        }

        Write-Log "  proxyAddresses final: $($newProxies -join '; ')"

        # Cast to string[] - Set-ADUser does not accept ArrayList
        [string[]]$newProxiesArray = $newProxies

        if ($PSCmdlet.ShouldProcess($label, "Update proxyAddresses")) {
            if ($currentProxies.Count -eq 0) {
                # Attribute may not exist yet - use -Add
                Set-ADUser -Identity $user.DistinguishedName -Add @{ proxyAddresses = $newProxiesArray } -ErrorAction Stop
            } else {
                # Attribute exists - use -Replace
                Set-ADUser -Identity $user.DistinguishedName -Replace @{ proxyAddresses = $newProxiesArray } -ErrorAction Stop
            }
            Write-Log "  proxyAddresses updated" -Level SUCCESS
        }

        # -- 3. Update mail attribute ---------------------------------------
        Write-Log "  mail: $($user.mail) -> $newEmail"

        if ($PSCmdlet.ShouldProcess($label, "Set mail to '$newEmail'")) {
            Set-ADUser -Identity $user.DistinguishedName -EmailAddress $newEmail -ErrorAction Stop
            Write-Log "  mail updated" -Level SUCCESS
        }

        $successCount++
        Write-Log "  Completed successfully" -Level SUCCESS
    }
    catch {
        $errorCount++
        Write-Log "  FAILED: $_" -Level ERROR
    }
}

# ===========================================================================
#  SUMMARY
# ===========================================================================
Write-Host ""
Write-Log "===================================================================="
Write-Log "MIGRATION COMPLETE"
Write-Log "  Processed   : $total"
Write-Log "  Successful  : $successCount" -Level SUCCESS
if ($errorCount -gt 0) {
    Write-Log "  Failed      : $errorCount" -Level ERROR
} else {
    Write-Log "  Failed      : 0"
}
Write-Log "  Log File    : $logFile"
Write-Log "  Backup File : $backupFile"
Write-Log "===================================================================="

# ===========================================================================
#  AZURE AD CONNECT - DELTA SYNC (optional, best-effort detection)
# ===========================================================================
if (-not $WhatIfPreference -and $successCount -gt 0) {
    Write-Log "Attempting to detect Azure AD Connect server(s)..."

    # -- Step 1: Discover all candidate servers from MSOL_ service accounts --
    $candidateServers = @()
    try {
        # Express installs create MSOL_<hex> service accounts whose Description
        # contains: "...running on computer <SERVERNAME> configured to..."
        $msolAccounts = @(
            Get-ADUser -Filter "SamAccountName -like 'MSOL_*'" -Properties Description -ErrorAction SilentlyContinue
        )
        foreach ($acct in $msolAccounts) {
            if ($acct.Description -match 'running on computer\s+(\S+)') {
                $server = $Matches[1].TrimEnd('.')
                if ($server -and $candidateServers -notcontains $server) {
                    $candidateServers += $server
                    Write-Log "  Found candidate AAD Connect server: $server (from $($acct.SamAccountName))"
                }
            }
        }
    } catch {
        # Detection is best-effort
    }

    if ($candidateServers.Count -eq 0) {
        Write-Log "Azure AD Connect server could not be auto-detected" -Level WARN
        Write-Log "If applicable, manually trigger a delta sync on your AAD Connect server" -Level WARN
    } else {
        # -- Step 2: Filter by computer account activity -------------------------
        # Read the domain's machine password rotation window (default 30 days)
        $maxPwdAgeDays = 30
        try {
            $domainPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
            if ($domainPolicy.MaxPasswordAge -and $domainPolicy.MaxPasswordAge.TotalDays -gt 0) {
                $maxPwdAgeDays = [int]$domainPolicy.MaxPasswordAge.TotalDays
            }
        } catch {
            Write-Log "  Could not read domain password policy, using default 30-day window" -Level WARN
        }
        $activityCutoff = (Get-Date).AddDays(-$maxPwdAgeDays)
        Write-Log "  Computer account activity cutoff: $($activityCutoff.ToString('yyyy-MM-dd')) ($maxPwdAgeDays day window)"

        $activeCandidates = @()
        foreach ($server in $candidateServers) {
            try {
                $computer = Get-ADComputer -Identity $server -Properties PasswordLastSet -ErrorAction Stop
                if ($computer.PasswordLastSet -and $computer.PasswordLastSet -ge $activityCutoff) {
                    $activeCandidates += $server
                    Write-Log "  $server - password last set $($computer.PasswordLastSet.ToString('yyyy-MM-dd')) - active" -Level SUCCESS
                } else {
                    $pwdDate = if ($computer.PasswordLastSet) { $computer.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'never' }
                    Write-Log "  $server - password last set $pwdDate - stale, skipping" -Level WARN
                }
            } catch {
                Write-Log "  $server - computer object not found in AD, skipping" -Level WARN
            }
        }

        if ($activeCandidates.Count -eq 0) {
            Write-Log "No active AAD Connect server computer accounts found (all stale or missing)" -Level WARN
            Write-Log "If applicable, manually trigger a delta sync on your AAD Connect server" -Level WARN
        } else {
            Write-Log "Checking connectivity and status for $($activeCandidates.Count) active candidate(s)..."

            $activeServer = $null

            foreach ($server in $activeCandidates) {
                Write-Log "  Checking $server..."

                # -- Step 3: Test WinRM connectivity --------------------------------
                try {
                    $wsmanResult = Test-WSMan -ComputerName $server -ErrorAction Stop
                    Write-Log "    WinRM: reachable" -Level SUCCESS
                } catch {
                    Write-Log "    WinRM: not reachable - $_" -Level WARN
                    continue
                }

                # -- Step 4: Test PowerShell remoting and check staging mode ---------
                try {
                    $syncStatus = Invoke-Command -ComputerName $server -ScriptBlock {
                        Import-Module ADSync -ErrorAction Stop
                        $scheduler = Get-ADSyncScheduler
                        [PSCustomObject]@{
                            StagingMode      = $scheduler.StagingModeEnabled
                            SyncEnabled      = $scheduler.SyncCycleEnabled
                            SchedulerRunning = $scheduler.SchedulerSuspended -eq $false
                        }
                    } -ErrorAction Stop

                    Write-Log "    Remoting: OK" -Level SUCCESS
                    Write-Log "    Staging mode: $($syncStatus.StagingMode)"
                    Write-Log "    Sync enabled: $($syncStatus.SyncEnabled)"

                    if ($syncStatus.StagingMode -eq $false) {
                        Write-Log "    $server is the ACTIVE sync server" -Level SUCCESS
                        $activeServer = $server
                        break
                    } else {
                        Write-Log "    $server is in STAGING mode - skipping" -Level WARN
                    }
                } catch {
                    Write-Log "    PowerShell remoting or ADSync query failed: $_" -Level WARN
                    continue
                }
            }

            # -- Step 5: Offer to trigger sync on the active server -----------------
            if ($activeServer) {
                $syncPrompt = Read-Host "Trigger a delta sync on active server '$activeServer'? (Y/N)"

                if ($syncPrompt -match '^[Yy]') {
                    try {
                        Invoke-Command -ComputerName $activeServer -ScriptBlock {
                            Import-Module ADSync -ErrorAction Stop
                            Start-ADSyncSyncCycle -PolicyType Delta
                        } -ErrorAction Stop
                        Write-Log "Delta sync initiated on $activeServer" -Level SUCCESS
                    } catch {
                        Write-Log "Failed to trigger delta sync: $_" -Level ERROR
                        Write-Log "Manually run on ${activeServer}: Start-ADSyncSyncCycle -PolicyType Delta" -Level WARN
                    }
                }
            } else {
                Write-Log "No active (non-staging) AAD Connect server found among candidates: $($candidateServers -join ', ')" -Level WARN
                Write-Log "If applicable, manually trigger a delta sync on your AAD Connect server" -Level WARN
            }
        }
    }
}

Write-Log "Script finished"
