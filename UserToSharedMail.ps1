<#
.SYNOPSIS
    Transfere the personnal mailbox of a deleted user to a shared mailbox.

.DESCRIPTION
    Transfere the personnal mailbox of a deleted user to a shared mailbox.
    The default behavior of Microsoft is to delete the shared mailbox linked to
    a user if the user is deleted. This scripts allows the shared mailbox to
    stay active even after the user is deleted.
    NOTE: the user must have been deleted in the last 30 days, during this time
    he is in a softdeleted state. After that the user no longer exists.

.PARAMETER Email
    The primary email address of the user to delete and migrate to a shared mailbox.
    This parameter is mandatory and must not be empty.

.PARAMETER ProxyFilter
    A string filter (regex) to select which proxy addresses from the user's mailbox to retain and apply to the new shared mailbox.
    Only proxy addresses matching this filter will be copied.

.PARAMETER FullAccessEmails
    An array of email addresses to grant Full Access mailbox permissions.

.PARAMETER ReviewerEmails
    An array of email addresses to grant Reviewer permissions on mailbox folders.

.PARAMETER MaxWaitMinutes
    The maximum number of minutes to wait for Azure AD user deletion to sync.
    Default is 30 minutes.

.PARAMETER DeleteAD
    Wether the Active Directory user should be deleted during the migration process.

.PARAMETER WhatIf
    Shows what actions would be performed without actually executing them.

.EXAMPLE
    .\UserToSharedMail.ps1 -Email user@example.com -FullAccessEmails admin@example.com,admin2@example.com -ReviewerEmails reviewer@example.com,reviewer2@example.com

    Deletes the specified user, creates a shared mailbox, assigns permissions, and restores the soft-deleted mailbox.

.NOTES
    This script run as a user.
    Requires ExchangeOnlineManagement and ActiveDirectory modules.
    Requires the following permissions:
    +------------------+-------+------------------------+---------------------------------------------------------------+
    | Component        | Scope | Role                   | Permissions                                                   |
    +------------------+-------+------------------------+---------------------------------------------------------------+
    | Active Directory | User  | Domain Administrator   | Delete a user from Active Directory                           |
    +------------------+-------+------------------------+---------------------------------------------------------------+
    | Entra            | User  | Exchange Administrator | Create shared mailbox and manage proxy addresses, permissions |
    +------------------+-------+------------------------+---------------------------------------------------------------+

.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchange/ (Exchange Online cmdlets)
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/ (Active Directory module)
#>
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$Email,
    [String]$ProxyFilter,
    [String[]]$FullAccessEmails,
    [String[]]$ReviewerEmails,
    [Int]$MaxWaitMinutes = 30,
    [Switch]$DeleteAD,
    [Switch]$WhatIf
)


#┌───────────────────────────┐
#│   PARAMETERS VALIDATION   │
#└───────────────────────────┘
function Test-ValidEmail {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Email
    )
    # Basic RFC 5322-compliant regex for email validation
    $emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return $Email -match $emailPattern
}
@($Email) + $FullAccessEmails + $ReviewerEmails | % {
    if ([string]::IsNullOrWhiteSpace($_)) { return }
    if(-not (Test-ValidEmail -Email $_)) {
        throw "Invalid email address: $_"
    }
}


#┌───────────────┐
#│   UTILITIES   │
#└───────────────┘
function Check-MailboxExists {
    param (
        [String]$Email,
        [Switch]$SoftDeleted
    )
    try {
        if($SoftDeleted) {
            $mailbox = Get-Mailbox -SoftDeletedMailbox -Identity $Email -ErrorAction Stop
        }
        else {
            $mailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
        }
        return $true
    }
    catch {
        return $false
    }
}
function Wait-UserDeletionExchange {
    param (
        [String]$Email,
        [Int]$MaxWaitMinutes = 30
    )
    
    Write-Host "Waiting for user deletion to sync (max $MaxWaitMinutes minutes)..."
    $startTime = Get-Date
    $maxWaitTime = $startTime.AddMinutes($MaxWaitMinutes)
    do {
        $userExists = $false
        try {
            Get-MailBox -Identity $Email -ErrorAction Stop
            $userExists = $true
        }
        catch {}
        
        if ($userExists) {
            Write-Host "User still exists. Waiting 30 seconds..."
            Start-Sleep -Seconds 30
        }
        else {
            Write-Host "User successfully deleted from all systems"
            return $true
        }
    } while ((Get-Date) -lt $maxWaitTime)
    Write-Host "Timeout waiting for Exchange user deletion."
    return $false
}
function RequestEXO-UserProxyAddresses {
    param (
        [String]$Email,
        [String]$ProxyFilter
    )
    $encodedEmail = [System.Web.HttpUtility]::UrlEncode($Email)
    $proxyAddresses = (Get-Mailbox -Identity $Email).EmailAddresses
    $proxyAddressesFiltered = $proxyAddresses | Where-Object { $_ -match $ProxyFilter }
    return $proxyAddressesFiltered
}
$whatIfPrefix = if($WhatIf) { "WHATIF: " } else { "" }

#┌─────────────┐
#│ CONNECTIONS │
#└─────────────┘
# Exchange Online connection
Write-Host "Connecting to ExchangeOnline"
Connect-ExchangeOnline


#┌───────────────────┐
#│    MAIN PROCESS   │
#└───────────────────┘
## 1. Request the addresses and apply filter
Write-Host "Retrieving proxy addresses:"
$proxyAddresses = @()
try {
    $proxyAddresses = RequestEXO-UserProxyAddresses -Email $Email -ProxyFilter $ProxyFilter
}
catch {
    throw "Failed to retrieve user proxy addresses: $($_.Exception.Message)" 
}
$proxyAddresses | ForEach-Object { Write-Host "    - $_"}

## 2. Delete the user from Active Directory. ##
if($DeleteAD) {
    $user = Get-ADUser -Filter {Mail -eq $Email}
    if ($user) {
        Write-Output $user
        if($WhatIf) {
            Write-Host "WHATIF: Would ask for the deletion of this user."
        }
        else {
            $confirmation = Read-Host "Are you sure you want to delete this user ? [Y/N]"
            if($confirmation -ne "Y") {
                throw "Operation cancelled."
            }
            else {
                Remove-ADUser -Identity $user -Confirm:$false
                Write-Host "AD user with email $Email has been deleted."
            }
        }
    }
    else {
        $confirmation = Read-Host "No user found with email $Email. Do you want to continue ? [Y/N]"
        if($confirmation -ne "Y") {
            throw "Operation cancelled."
        }
    }
}

## 3. Wait for the user to be deleted from Azure. ##
Write-Host "$($whatIfPrefix)Waiting for the deletion of the of the user in Exchange"
if(-not $WhatIf) {
    Wait-UserDeletionExchange -Email $Email -MaxWaitMinutes $MaxWaitMinutes
}

## 4. Create a new shared mailbox with the old mail. ##
$nameParts = ($Email -split "@")[0] -split "\."
$firstName = $nameParts[0]
$lastName = $nameParts[1]
Write-Host "$($whatIfPrefix)Create a new shared mailbox: $firstName $lastName $Email"
if(-not $WhatIf) {
    New-Mailbox -Shared -Name "$firstName $lastName" -DisplayName "$($firstName) $($lastName)" -Alias "$($firstName)$($lastName)" -PrimarySmtpAddress $Email
}

## 5. Add proxy addresses to the new shared mailbox. ##
Write-Host "$($whatIfPrefix)Add proxy addresses to $Email"
if(-not $WhatIf) {
    Set-Mailbox -Identity $Email -EmailAddresses @{add=$proxyAddresses}
}

## 6. Gives the permission to new owners. ##
### Giving full acess is easy
if($FullAccessEmails -and $FullAccessEmails.Count -gt 0) {
    Write-Host "$($whatIfPrefix)Giving full access permission to:"
    $FullAccessEmails | ForEach-Object {
        Write-Host "  $($whatIfPrefix)$_"
        if(-not $WhatIf) {
            Add-MailboxPermission -Identity $Email -User $_ -AccessRights FullAccess
        }
    }
}

### For reviewer access we need to give folder based permission
if($ReviewerEmails -and $ReviewerEmails.Count -gt 0) {
    Write-Host "$($whatIfPrefix)Giving reviewer permission to:"
    $ReviewerEmails | Where-Object { $_ -ne $null } | ForEach-Object {
        Write-Host "  $($whatIfPrefix)$_"
        $folders = Get-MailboxFolderStatistics -Identity $Email | Where-Object {$_.FolderType -ne "SearchFolder"}
        foreach($folder in $folders) {
            $folderPath = $folder.FolderPath.Replace("/", "\")  # Format correctly
            $identity = "$Email`\:$folderPath" 
            try {
                Write-Host "    $($whatIfPrefix)$identity"
                if(-not $WhatIf) {
                    Add-MailboxFolderPermission -Identity $identity -User $_ -AccessRights Reviewer
                }
            }
            catch {
                Write-Warning "Failed to set permission on $identity for $_"
            }
        }
    }
}

## 7. Get the GUID from the old mailbox and the new mailbox (even though they have the same name you can get both, see the script). ##
Write-Host "$($whatIfPrefix)Get the GUID of the soft deleted and the new shared mailbox with address: $Email"
if(-not $WhatIf) {
    $deletedEmail = Get-Mailbox -SoftDeletedMailbox -Identity $Email
    $newEmail = Get-Mailbox -Identity $Email
}

## 8. Restore the old mailbox to the new one. ##
Write-Host "$($whatIfPrefix)Restore soft deleted mailbox $Email to new mailbox $Email"
if(-not $WhatIf) {
    New-MailboxRestoreRequest -SourceMailbox $deletedEmail.ExchangeGuid -TargetMailbox $newEmail.ExchangeGuid -AllowLegacyDNMismatch
}
