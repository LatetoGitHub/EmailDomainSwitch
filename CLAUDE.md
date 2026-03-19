# EmailDomainSwitch

PowerShell scripts for migrating Microsoft 365 email domains -- both on-premises Active Directory users (synced via Azure AD Connect) and cloud-only Exchange Online recipients.

## Scripts

### Switch-EmailDomain.ps1

Updates on-premises AD users during a domain migration:

- **UPN**: Changes the UserPrincipalName to the new domain
- **proxyAddresses**: Sets the new domain as primary SMTP, demotes the old primary to alias, ensures the old UPN-based address exists as an alias
- **mail attribute**: Updated to match the new primary SMTP address
- **Azure AD Connect**: Attempts to detect the active (non-staging) AAD Connect server via MSOL_ service accounts, validates computer account activity, checks WinRM/PS remoting connectivity, and offers to trigger a delta sync

Parameters: `-NewDomain` (required), `-OUName` (required), `-IncludeDisabledUsers`, `-WhatIf`

Requires: `ActiveDirectory` PowerShell module

### Switch-CloudEmailDomain.ps1

Updates cloud-only (non-directory-synced) Exchange Online recipients in four phases:

1. **User mailboxes + shared mailboxes** (Global Admins excluded)
2. **Distribution groups** (MailUniversalDistributionGroup only)
3. **Microsoft 365 (Unified) groups**
4. **Mail-enabled security groups** (MailUniversalSecurityGroup)

Each phase lists affected objects, prompts for confirmation, then applies changes. Global Admin role members are detected via Microsoft Graph (`Microsoft.Graph.Identity.DirectoryManagement` + `Microsoft.Graph.Users` modules) and excluded from mailbox changes.

Parameters: `-NewDomain` (required), `-WhatIf`

Requires: `ExchangeOnlineManagement`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Users`

### Repair-MissingAliases.ps1

Remediation script for users who were migrated by Switch-EmailDomain.ps1 but are missing an alias for their old email address. Reads the pre-migration backup CSV to determine each user's old UPN, checks current proxyAddresses, and adds the missing `smtp:` alias.

Parameters: `-BackupCsvPath` (required), `-WhatIf`

Requires: `ActiveDirectory` PowerShell module

## Shared behavior

- All scripts support `-WhatIf` to preview changes without applying them
- All changes are logged to timestamped log files in the script's directory
- Pre-migration backups are exported to CSV before any modifications
- Old primary email addresses are always preserved as aliases
- Existing aliases for the new domain are promoted rather than duplicated
- Non-SMTP proxy entries (X500, SIP, SPO) are preserved

## Output files

Generated at runtime (excluded from git via .gitignore):

- `DomainMigration_<timestamp>.log` -- AD script log
- `PreMigrationBackup_<timestamp>.csv` -- AD script backup
- `CloudDomainMigration_<timestamp>.log` -- Cloud script log
- `CloudPreMigrationBackup_<timestamp>.csv` -- Cloud script backup
- `AliasRemediation_<timestamp>.log` -- Repair script log

## Known issues addressed

- `Set-ADUser` does not accept `System.Collections.ArrayList` for proxyAddresses -- cast to `[string[]]` before passing
- `Write-Log` uses `Out-File -WhatIf:$false` to prevent `-WhatIf` propagation from suppressing log file writes
- `Get-MgDirectoryRoleMember` returns `DirectoryObject` references where `AdditionalProperties["userPrincipalName"]` is unreliable across Graph SDK versions -- resolved by fetching each member via `Get-MgUser -UserId`
- Existing Graph sessions may lack required scopes -- script checks scopes and reconnects if `RoleManagement.Read.Directory` or `User.Read.All` is missing

## Code style

- ASCII-only characters in all code and output (no Unicode box-drawing, arrows, or symbols)
- PowerShell 5.1 compatible
