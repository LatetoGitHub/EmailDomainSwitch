<#
.SYNOPSIS
    Updates primary email addresses for cloud-only mailboxes, distribution groups,
    mail-enabled security groups, and Microsoft 365 groups during a domain migration.

.DESCRIPTION
    Processes the following cloud-only (non-directory-synced) recipient types in order:
      1. User mailboxes and shared mailboxes (Global Admins excluded)
      2. Distribution groups
      3. Microsoft 365 (Unified) groups
      4. Mail-enabled security groups

    For each category the script lists all objects that will be updated, prompts for
    confirmation, then applies the changes. The previous primary email address is kept
    as an alias. All changes are logged to a timestamped log file.

    Global Admin role members are automatically detected via Microsoft Graph and
    excluded from all mailbox changes.

.PARAMETER NewDomain
    The new email domain (e.g., "contoso.com"). Do not include the @ symbol.

.PARAMETER WhatIf
    Preview all changes without applying them.

.EXAMPLE
    .\Switch-CloudEmailDomain.ps1 -NewDomain "contoso.com"

.EXAMPLE
    .\Switch-CloudEmailDomain.ps1 -NewDomain "contoso.com" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "New email domain (e.g., contoso.com)")]
    [ValidateNotNullOrEmpty()]
    [string]$NewDomain
)

#Requires -Modules ExchangeOnlineManagement

# ============================================================================
#  LOGGING
# ============================================================================
$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile    = Join-Path $scriptDir "CloudDomainMigration_$timestamp.log"
$backupFile = Join-Path $scriptDir "CloudPreMigrationBackup_$timestamp.csv"

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

# ============================================================================
#  HELPER: Build updated email address list
# ============================================================================
function Build-UpdatedEmailAddresses {
    param(
        [string[]]$CurrentAddresses,
        [string]$NewDomain
    )

    $currentPrimary = $CurrentAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1
    if (-not $currentPrimary) {
        return @{ Changed = $false; Reason = "No primary SMTP address found" }
    }

    $primaryEmail  = ($currentPrimary -split ":", 2)[1]
    $localPart     = ($primaryEmail -split "@")[0]
    $currentDomain = ($primaryEmail -split "@")[1]

    if ($currentDomain -ieq $NewDomain) {
        return @{ Changed = $false; Reason = "Already on new domain" }
    }

    $newPrimaryEmail = "$localPart@$NewDomain"
    $newAddresses = [System.Collections.ArrayList]@()

    # Check if new domain address already exists as an alias
    $existingNewEntry = $CurrentAddresses | Where-Object {
        ($_ -split ":", 2)[1] -ieq $newPrimaryEmail
    } | Select-Object -First 1

    foreach ($addr in $CurrentAddresses) {
        if ($existingNewEntry -and $addr -ceq $existingNewEntry) {
            # Promote existing alias to primary
            [void]$newAddresses.Add("SMTP:$newPrimaryEmail")
        }
        elseif ($addr -cmatch '^SMTP:') {
            # Demote current primary to alias
            $emailPart = ($addr -split ":", 2)[1]
            [void]$newAddresses.Add("smtp:$emailPart")
        }
        else {
            # Keep everything else as-is (aliases, X500, SIP, SPO, etc.)
            [void]$newAddresses.Add($addr)
        }
    }

    if (-not $existingNewEntry) {
        [void]$newAddresses.Add("SMTP:$newPrimaryEmail")
    }

    # Ensure old primary exists as alias
    $oldAlias = "smtp:$primaryEmail"
    $hasOldAlias = $newAddresses | Where-Object { $_ -ieq $oldAlias }
    if (-not $hasOldAlias) {
        [void]$newAddresses.Add($oldAlias)
    }

    return @{
        Changed      = $true
        NewAddresses = [string[]]$newAddresses
        OldPrimary   = $primaryEmail
        NewPrimary   = $newPrimaryEmail
    }
}

# ============================================================================
#  INITIALISATION
# ============================================================================
$NewDomain = $NewDomain.TrimStart("@").Trim()

Write-Log "Cloud Email Domain Migration Script Started"
Write-Log "New Domain  : $NewDomain"
Write-Log "WhatIf Mode : $WhatIfPreference"
Write-Log "Log File    : $logFile"

# -- Check / establish Exchange Online connection ----------------------------
Write-Log "Checking Exchange Online connection..."
try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Log "Already connected to Exchange Online" -Level SUCCESS
} catch {
    Write-Log "Not connected. Attempting to connect to Exchange Online..."
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected to Exchange Online" -Level SUCCESS
    } catch {
        Write-Log "Failed to connect to Exchange Online: $_" -Level ERROR
        exit 1
    }
}

# ============================================================================
#  GLOBAL ADMIN EXCLUSION LIST
# ============================================================================
Write-Log "Retrieving Global Admin role members for exclusion..."

$globalAdminEmails = @()
$exclusionLoaded   = $false

try {
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    $requiredScopes = @("RoleManagement.Read.Directory", "User.Read.All")
    $ctx = Get-MgContext -ErrorAction SilentlyContinue

    # Reconnect if no session or if required scopes are missing
    $needsConnect = $true
    if ($ctx -and $ctx.Scopes) {
        $missingScopes = $requiredScopes | Where-Object { $ctx.Scopes -notcontains $_ }
        if ($missingScopes.Count -eq 0) {
            $needsConnect = $false
            Write-Log "Existing Microsoft Graph session has required scopes"
        } else {
            Write-Log "Existing Graph session missing scopes: $($missingScopes -join ', '). Reconnecting..." -Level WARN
        }
    }
    if ($needsConnect) {
        Write-Log "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
    }

    # Global Administrator role template ID
    $role = Get-MgDirectoryRole -Filter "roleTemplateId eq '62e90394-69f5-4237-9190-012177145e10'" -ErrorAction Stop
    if ($role) {
        $memberRefs = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
        Write-Log "  Found $($memberRefs.Count) role member(s), resolving identities..."
        foreach ($ref in $memberRefs) {
            try {
                # Fetch the full user object by ID to reliably get the UPN
                $mgUser = Get-MgUser -UserId $ref.Id -Property UserPrincipalName -ErrorAction Stop
                if ($mgUser.UserPrincipalName) {
                    $globalAdminEmails += $mgUser.UserPrincipalName.ToLower()
                }
            } catch {
                # Log the actual error -- could be a service principal, or a permissions issue
                $errMsg = $_.Exception.Message
                if ($errMsg -match 'Resource.*not found' -or $errMsg -match 'does not exist') {
                    Write-Log "  Skipping non-user role member (Id: $($ref.Id))" -Level WARN
                } else {
                    Write-Log "  Failed to resolve member $($ref.Id): $errMsg" -Level ERROR
                }
            }
        }
    }

    $exclusionLoaded = $true
    Write-Log "Global Admin exclusion list: $($globalAdminEmails.Count) member(s)" -Level SUCCESS
    foreach ($ga in $globalAdminEmails) {
        Write-Log "  Excluding: $ga"
    }
} catch {
    Write-Log "Failed to retrieve Global Admin members: $_" -Level WARN
    Write-Log "The Microsoft.Graph.Identity.DirectoryManagement and Microsoft.Graph.Users modules are required" -Level WARN
    $cont = Read-Host "Continue WITHOUT excluding Global Admins? (Y/N)"
    if ($cont -notmatch '^[Yy]') {
        Write-Log "Script cancelled by user" -Level WARN
        exit 0
    }
}

# ============================================================================
#  PHASE 1: CLOUD-ONLY USER MAILBOXES + SHARED MAILBOXES
# ============================================================================
Write-Log "===================================================================="
Write-Log "PHASE 1: Cloud-Only Mailboxes (User + Shared)"
Write-Log "===================================================================="
Write-Log "Retrieving cloud-only mailboxes..."

try {
    $allMailboxes = @(
        Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox |
            Where-Object { $_.IsDirSynced -eq $false }
    )
    Write-Log "Found $($allMailboxes.Count) cloud-only mailbox(es)"
} catch {
    Write-Log "Failed to retrieve mailboxes: $_" -Level ERROR
    exit 1
}

# Exclude Global Admins
$gaExcludedCount = 0
if ($globalAdminEmails.Count -gt 0) {
    $beforeCount  = $allMailboxes.Count
    $allMailboxes = @($allMailboxes | Where-Object {
        $_.PrimarySmtpAddress.ToString().ToLower() -notin $globalAdminEmails -and
        $_.UserPrincipalName.ToLower() -notin $globalAdminEmails
    })
    $gaExcludedCount = $beforeCount - $allMailboxes.Count
    if ($gaExcludedCount -gt 0) {
        Write-Log "Excluded $gaExcludedCount Global Admin mailbox(es)" -Level WARN
    }
}

# Filter to those needing update
$mailboxesToUpdate = @()
foreach ($mbx in $allMailboxes) {
    $result = Build-UpdatedEmailAddresses -CurrentAddresses $mbx.EmailAddresses -NewDomain $NewDomain
    if ($result.Changed) {
        $mailboxesToUpdate += [PSCustomObject]@{
            Identity           = $mbx.Identity
            DisplayName        = $mbx.DisplayName
            Type               = $mbx.RecipientTypeDetails
            CurrentPrimary     = $result.OldPrimary
            NewPrimary         = $result.NewPrimary
            NewAddresses       = $result.NewAddresses
            CurrentAddresses   = $mbx.EmailAddresses
        }
    }
}

$mailboxSuccessCount = 0
$mailboxErrorCount   = 0

if ($mailboxesToUpdate.Count -eq 0) {
    Write-Log "No mailboxes need updating" -Level WARN
} else {
    Write-Log "$($mailboxesToUpdate.Count) mailbox(es) to update:"
    Write-Host ""
    $mailboxesToUpdate |
        Select-Object @{N="#"; E={[array]::IndexOf($mailboxesToUpdate, $_) + 1}},
                      DisplayName, Type, CurrentPrimary, NewPrimary |
        Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ }

    # Export backup for mailboxes
    $mailboxesToUpdate | Select-Object Identity, DisplayName, Type, CurrentPrimary,
        @{N="EmailAddresses"; E={($_.CurrentAddresses -join ";")}} |
        Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8 -Append
    Write-Log "Backup exported to: $backupFile"

    if (-not $WhatIfPreference) {
        $confirm = Read-Host "Proceed with updating $($mailboxesToUpdate.Count) mailbox(es)? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Mailbox updates skipped by user" -Level WARN
        } else {
            $counter = 0
            foreach ($mbx in $mailboxesToUpdate) {
                $counter++
                Write-Log "[$counter/$($mailboxesToUpdate.Count)] $($mbx.DisplayName) ($($mbx.Type))"
                Write-Log "  Primary: $($mbx.CurrentPrimary) -> $($mbx.NewPrimary)"
                try {
                    if ($PSCmdlet.ShouldProcess($mbx.DisplayName, "Set primary to '$($mbx.NewPrimary)'")) {
                        Set-Mailbox -Identity $mbx.Identity -EmailAddresses $mbx.NewAddresses -ErrorAction Stop
                        Write-Log "  Updated successfully" -Level SUCCESS
                        $mailboxSuccessCount++
                    }
                } catch {
                    Write-Log "  FAILED: $_" -Level ERROR
                    $mailboxErrorCount++
                }
            }
        }
    } else {
        Write-Log "[WHATIF] Would update $($mailboxesToUpdate.Count) mailbox(es)" -Level WARN
    }
}

# ============================================================================
#  PHASE 2: CLOUD-ONLY DISTRIBUTION GROUPS
# ============================================================================
Write-Log "===================================================================="
Write-Log "PHASE 2: Cloud-Only Distribution Groups"
Write-Log "===================================================================="
Write-Log "Retrieving cloud-only distribution groups..."

try {
    $allDLs = @(
        Get-DistributionGroup -ResultSize Unlimited |
            Where-Object {
                $_.IsDirSynced -eq $false -and
                $_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup'
            }
    )
    Write-Log "Found $($allDLs.Count) cloud-only distribution group(s)"
} catch {
    Write-Log "Failed to retrieve distribution groups: $_" -Level ERROR
    $allDLs = @()
}

$dlsToUpdate = @()
foreach ($dl in $allDLs) {
    $result = Build-UpdatedEmailAddresses -CurrentAddresses $dl.EmailAddresses -NewDomain $NewDomain
    if ($result.Changed) {
        $dlsToUpdate += [PSCustomObject]@{
            Identity         = $dl.Identity
            DisplayName      = $dl.DisplayName
            CurrentPrimary   = $result.OldPrimary
            NewPrimary       = $result.NewPrimary
            NewAddresses     = $result.NewAddresses
            CurrentAddresses = $dl.EmailAddresses
        }
    }
}

$dlSuccessCount = 0
$dlErrorCount   = 0

if ($dlsToUpdate.Count -eq 0) {
    Write-Log "No distribution groups need updating" -Level WARN
} else {
    Write-Log "$($dlsToUpdate.Count) distribution group(s) to update:"
    Write-Host ""
    $dlsToUpdate |
        Select-Object @{N="#"; E={[array]::IndexOf($dlsToUpdate, $_) + 1}},
                      DisplayName, CurrentPrimary, NewPrimary |
        Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ }

    # Append backup
    $dlsToUpdate | Select-Object Identity, DisplayName,
        @{N="Type"; E={"DistributionGroup"}}, CurrentPrimary,
        @{N="EmailAddresses"; E={($_.CurrentAddresses -join ";")}} |
        Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8 -Append
    Write-Log "Backup appended to: $backupFile"

    if (-not $WhatIfPreference) {
        $confirm = Read-Host "Proceed with updating $($dlsToUpdate.Count) distribution group(s)? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Distribution group updates skipped by user" -Level WARN
        } else {
            $counter = 0
            foreach ($dl in $dlsToUpdate) {
                $counter++
                Write-Log "[$counter/$($dlsToUpdate.Count)] $($dl.DisplayName)"
                Write-Log "  Primary: $($dl.CurrentPrimary) -> $($dl.NewPrimary)"
                try {
                    if ($PSCmdlet.ShouldProcess($dl.DisplayName, "Set primary to '$($dl.NewPrimary)'")) {
                        Set-DistributionGroup -Identity $dl.Identity -EmailAddresses $dl.NewAddresses -ErrorAction Stop
                        Write-Log "  Updated successfully" -Level SUCCESS
                        $dlSuccessCount++
                    }
                } catch {
                    Write-Log "  FAILED: $_" -Level ERROR
                    $dlErrorCount++
                }
            }
        }
    } else {
        Write-Log "[WHATIF] Would update $($dlsToUpdate.Count) distribution group(s)" -Level WARN
    }
}

# ============================================================================
#  PHASE 3: CLOUD-ONLY MICROSOFT 365 GROUPS
# ============================================================================
Write-Log "===================================================================="
Write-Log "PHASE 3: Cloud-Only Microsoft 365 Groups"
Write-Log "===================================================================="
Write-Log "Retrieving cloud-only Microsoft 365 groups..."

try {
    $allM365Groups = @(
        Get-UnifiedGroup -ResultSize Unlimited |
            Where-Object { $_.IsDirSynced -eq $false }
    )
    Write-Log "Found $($allM365Groups.Count) cloud-only Microsoft 365 group(s)"
} catch {
    Write-Log "Failed to retrieve Microsoft 365 groups: $_" -Level ERROR
    $allM365Groups = @()
}

$m365ToUpdate = @()
foreach ($grp in $allM365Groups) {
    $result = Build-UpdatedEmailAddresses -CurrentAddresses $grp.EmailAddresses -NewDomain $NewDomain
    if ($result.Changed) {
        $m365ToUpdate += [PSCustomObject]@{
            Identity         = $grp.Identity
            DisplayName      = $grp.DisplayName
            CurrentPrimary   = $result.OldPrimary
            NewPrimary       = $result.NewPrimary
            NewAddresses     = $result.NewAddresses
            CurrentAddresses = $grp.EmailAddresses
        }
    }
}

$m365SuccessCount = 0
$m365ErrorCount   = 0

if ($m365ToUpdate.Count -eq 0) {
    Write-Log "No Microsoft 365 groups need updating" -Level WARN
} else {
    Write-Log "$($m365ToUpdate.Count) Microsoft 365 group(s) to update:"
    Write-Host ""
    $m365ToUpdate |
        Select-Object @{N="#"; E={[array]::IndexOf($m365ToUpdate, $_) + 1}},
                      DisplayName, CurrentPrimary, NewPrimary |
        Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ }

    # Append backup
    $m365ToUpdate | Select-Object Identity, DisplayName,
        @{N="Type"; E={"UnifiedGroup"}}, CurrentPrimary,
        @{N="EmailAddresses"; E={($_.CurrentAddresses -join ";")}} |
        Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8 -Append
    Write-Log "Backup appended to: $backupFile"

    if (-not $WhatIfPreference) {
        $confirm = Read-Host "Proceed with updating $($m365ToUpdate.Count) Microsoft 365 group(s)? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Microsoft 365 group updates skipped by user" -Level WARN
        } else {
            $counter = 0
            foreach ($grp in $m365ToUpdate) {
                $counter++
                Write-Log "[$counter/$($m365ToUpdate.Count)] $($grp.DisplayName)"
                Write-Log "  Primary: $($grp.CurrentPrimary) -> $($grp.NewPrimary)"
                try {
                    if ($PSCmdlet.ShouldProcess($grp.DisplayName, "Set primary to '$($grp.NewPrimary)'")) {
                        Set-UnifiedGroup -Identity $grp.Identity -EmailAddresses $grp.NewAddresses -ErrorAction Stop
                        Write-Log "  Updated successfully" -Level SUCCESS
                        $m365SuccessCount++
                    }
                } catch {
                    Write-Log "  FAILED: $_" -Level ERROR
                    $m365ErrorCount++
                }
            }
        }
    } else {
        Write-Log "[WHATIF] Would update $($m365ToUpdate.Count) Microsoft 365 group(s)" -Level WARN
    }
}

# ============================================================================
#  PHASE 4: CLOUD-ONLY MAIL-ENABLED SECURITY GROUPS
# ============================================================================
Write-Log "===================================================================="
Write-Log "PHASE 4: Cloud-Only Mail-Enabled Security Groups"
Write-Log "===================================================================="
Write-Log "Retrieving cloud-only mail-enabled security groups..."

try {
    $allMESGs = @(
        Get-DistributionGroup -ResultSize Unlimited |
            Where-Object {
                $_.IsDirSynced -eq $false -and
                $_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup'
            }
    )
    Write-Log "Found $($allMESGs.Count) cloud-only mail-enabled security group(s)"
} catch {
    Write-Log "Failed to retrieve mail-enabled security groups: $_" -Level ERROR
    $allMESGs = @()
}

$mesgsToUpdate = @()
foreach ($sg in $allMESGs) {
    $result = Build-UpdatedEmailAddresses -CurrentAddresses $sg.EmailAddresses -NewDomain $NewDomain
    if ($result.Changed) {
        $mesgsToUpdate += [PSCustomObject]@{
            Identity         = $sg.Identity
            DisplayName      = $sg.DisplayName
            CurrentPrimary   = $result.OldPrimary
            NewPrimary       = $result.NewPrimary
            NewAddresses     = $result.NewAddresses
            CurrentAddresses = $sg.EmailAddresses
        }
    }
}

$mesgSuccessCount = 0
$mesgErrorCount   = 0

if ($mesgsToUpdate.Count -eq 0) {
    Write-Log "No mail-enabled security groups need updating" -Level WARN
} else {
    Write-Log "$($mesgsToUpdate.Count) mail-enabled security group(s) to update:"
    Write-Host ""
    $mesgsToUpdate |
        Select-Object @{N="#"; E={[array]::IndexOf($mesgsToUpdate, $_) + 1}},
                      DisplayName, CurrentPrimary, NewPrimary |
        Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ }

    # Append backup
    $mesgsToUpdate | Select-Object Identity, DisplayName,
        @{N="Type"; E={"MailEnabledSecurityGroup"}}, CurrentPrimary,
        @{N="EmailAddresses"; E={($_.CurrentAddresses -join ";")}} |
        Export-Csv -Path $backupFile -NoTypeInformation -Encoding UTF8 -Append
    Write-Log "Backup appended to: $backupFile"

    if (-not $WhatIfPreference) {
        $confirm = Read-Host "Proceed with updating $($mesgsToUpdate.Count) mail-enabled security group(s)? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Mail-enabled security group updates skipped by user" -Level WARN
        } else {
            $counter = 0
            foreach ($sg in $mesgsToUpdate) {
                $counter++
                Write-Log "[$counter/$($mesgsToUpdate.Count)] $($sg.DisplayName)"
                Write-Log "  Primary: $($sg.CurrentPrimary) -> $($sg.NewPrimary)"
                try {
                    if ($PSCmdlet.ShouldProcess($sg.DisplayName, "Set primary to '$($sg.NewPrimary)'")) {
                        Set-DistributionGroup -Identity $sg.Identity -EmailAddresses $sg.NewAddresses -ErrorAction Stop
                        Write-Log "  Updated successfully" -Level SUCCESS
                        $mesgSuccessCount++
                    }
                } catch {
                    Write-Log "  FAILED: $_" -Level ERROR
                    $mesgErrorCount++
                }
            }
        }
    } else {
        Write-Log "[WHATIF] Would update $($mesgsToUpdate.Count) mail-enabled security group(s)" -Level WARN
    }
}

# ============================================================================
#  SUMMARY
# ============================================================================
$totalSuccess = $mailboxSuccessCount + $dlSuccessCount + $m365SuccessCount + $mesgSuccessCount
$totalErrors  = $mailboxErrorCount + $dlErrorCount + $m365ErrorCount + $mesgErrorCount

Write-Host ""
Write-Log "===================================================================="
Write-Log "MIGRATION COMPLETE"
Write-Log "--------------------------------------------------------------------"
Write-Log "  Mailboxes              : $mailboxSuccessCount succeeded, $mailboxErrorCount failed (of $($mailboxesToUpdate.Count))"
if ($gaExcludedCount -gt 0) {
    Write-Log "  Global Admins Excluded : $gaExcludedCount"
}
Write-Log "  Distribution Groups    : $dlSuccessCount succeeded, $dlErrorCount failed (of $($dlsToUpdate.Count))"
Write-Log "  Microsoft 365 Groups   : $m365SuccessCount succeeded, $m365ErrorCount failed (of $($m365ToUpdate.Count))"
Write-Log "  Mail-Enabled Sec Groups: $mesgSuccessCount succeeded, $mesgErrorCount failed (of $($mesgsToUpdate.Count))"
Write-Log "--------------------------------------------------------------------"
Write-Log "  Total Successful       : $totalSuccess" -Level SUCCESS
if ($totalErrors -gt 0) {
    Write-Log "  Total Failed           : $totalErrors" -Level ERROR
} else {
    Write-Log "  Total Failed           : 0"
}
Write-Log "  Log File               : $logFile"
Write-Log "  Backup File            : $backupFile"
Write-Log "===================================================================="
Write-Log "Script finished"
