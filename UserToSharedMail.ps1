<#
.SYNOPSIS
    Transfer the personal mailbox of a deleted user to a shared mailbox.

.DESCRIPTION
    Transfer the personal mailbox of a deleted user to a shared mailbox.
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
    Defaults to ".*" (include all proxy addresses).

.PARAMETER FullAccessEmails
    An array of email addresses to grant Full Access mailbox permissions.

.PARAMETER ReviewerEmails
    An array of email addresses to grant Reviewer permissions on mailbox folders.

.PARAMETER MaxWaitMinutes
    The maximum number of minutes to wait for Azure AD user deletion to sync.
    Default is 30 minutes.

.PARAMETER DeleteAD
    If true the Active Directory user should be deleted during the migration process.

.PARAMETER Archive
    If true, the script will archive the mailbox by creating a shared mailbox 
    and restoring the deleted mailbox contents into it.

.PARAMETER RedirectEmail
    An email address to which future mail for the mailbox will be forwarded.
    Can be internal or external depending on the value of -RedirectExternal.

.PARAMETER RedirectExternal
    If true, treats the -RedirectEmail as an external email address 
    and sets ForwardingSMTPAddress instead of ForwardingAddress.

.PARAMETER DeliverToMailboxAndForward
    If true, retains copies of forwarded emails in the original shared mailbox.

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
[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$Email,
    [String]$ProxyFilter = ".*",
    [String[]]$FullAccessEmails,
    [String[]]$ReviewerEmails,
    [Int]$MaxWaitMinutes = 30,
    [Bool]$DeleteAD = $false,
    [Bool]$Archive = $true,
    [String]$RedirectEmail ,
    [Bool]$RedirectExternal = $false,
    [Bool]$DeliverToMailboxAndForward = $true
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
@($Email) + $FullAccessEmails + $ReviewerEmails + @($RedirectEmail) | % {
    if ([string]::IsNullOrWhiteSpace($_)) { return }
    if(-not (Test-ValidEmail -Email $_)) {
        throw "Invalid email address: $_"
    }
}
if($Archive) {
    try {
        Get-Mailbox -ResultSize 1 -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Not connected to Exchange Online"
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

# Archive utilities
function Get-MailboxProxyAddresses {
    param (
        [String]$Email
    )
    $encodedEmail = [System.Web.HttpUtility]::UrlEncode($Email)
    try {
        return (Get-Mailbox -Identity $Email).EmailAddresses
    }
    catch {
        return (Get-Mailbox -SoftDeletedMailbox -Identity $Email).EmailAddresses
    }
}
function Delete-ADUserFromEmail {
    param (
        [String]$Email
    )
    $user = Get-ADUser -Filter {Mail -eq $Email}
    if($user -eq $null) {
        Write-Warning "No user found in AD with email $Email."
        return
    }
    $children = Get-ADObject -Filter 'ObjectClass -ne "user"' -SearchBase (Get-ADUser $user).DistinguishedName
    if($children -and $children.Count -gt 0) {
        Write-Host "User with email $Email has the following children in Active Directory:"
        $children | ForEach-Object { Write-Host "    - $_" }
    }
    if($PSCmdlet.ShouldProcess("$($user.SamAccountName)", "Remove AD user")) {
        Remove-ADUser -Identity $user -Confirm:$false
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
            Get-Mailbox -Identity $Email -ErrorAction Stop
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
function Add-FullAccessPermission {
    param (
        [String]$Email,
        [String[]]$FullAccessEmails
    )
    if($FullAccessEmails -and $FullAccessEmails.Count -gt 0) {
        $FullAccessEmails | ForEach-Object {
            if($PSCmdlet.ShouldProcess("$_", "Grant $Email FullAccess permission")) {
                Add-MailboxPermission -Identity $Email `
                                      -User $_ `
                                      -AccessRights FullAccess `
                                      -Confirm:$false
            }
        }
    }
}
function Add-ReviewerPermission {
    param (
        [String]$Email,
        [String[]]$ReviewerEmails
    )
    # For reviewer access we need to give folder based permission
    if($ReviewerEmails -and $ReviewerEmails.Count -gt 0) {
        $ReviewerEmails | ForEach-Object {
            $folders = Get-MailboxFolderStatistics -Identity $Email | Where-Object {$_.FolderType -ne "SearchFolder"}
            foreach($folder in $folders) {
                $folderPath = $folder.FolderPath.Replace("/", "\")
                $identity = "$Email`\:$folderPath" 
                if($PSCmdlet.ShouldProcess("$_", "Grant $Email Reviewer permissions")) {
                    try {
                        Add-MailboxFolderPermission -Identity $identity `
                                                    -User $_ `
                                                    -AccessRights Reviewer `
                                                    -Confirm:$false
                    }
                    catch {
                        Write-Warning "Failed to set permission on $identity for $_"
                    }
                }
            }
        }
    }
}
function Restore-OldToNewMailbox {
    param (
        [String]$Email
    )
    if($PSCmdlet.ShouldProcess("Mailbox $Email", "Restore soft deleted mailbox to new shared email")) {
        # Get the GUID from the old mailbox and the new mailbox.
        $deletedEmail = Get-Mailbox -SoftDeletedMailbox -Identity $Email
        $newEmail = Get-Mailbox -Identity $Email

        # Restore the old mailbox to the new one.
        New-MailboxRestoreRequest -SourceMailbox $deletedEmail.ExchangeGuid `
                                  -TargetMailbox $newEmail.ExchangeGuid `
                                  -AllowLegacyDNMismatch `
                                  -Confirm:$false
    }
}
function CreateAndRestore-OldToNewMailbox {
    param (
        [String]$Email,
        [String]$ProxyFilter,
        [String[]]$FullAccessEmails,
        [String[]]$ReviewerEmails
    )
    # Retrieve proxy addresses.
    $proxyAddresses = Get-MailboxProxyAddresses -Email $Email | Where-Object { $_ -match $ProxyFilter }

    # Create a new shared mailbox with the old mail.
    $nameParts = ($Email -split "@")[0] -split "\."
    $firstName = $nameParts[0]
    $lastName = $nameParts[1]
    if($PSCmdlet.ShouldProcess("Mailbox $Email", "Create new shared mailbox")) {
        New-Mailbox -Shared `
                    -Name "$firstName $lastName" `
                    -DisplayName "$firstName $lastName" `
                    -Alias "$firstName$lastName" `
                    -PrimarySmtpAddress $Email `
                    -Confirm:$false
    }

    # Add proxy addresses to the new shared mailbox.
    if($PSCmdlet.ShouldProcess("Mailbox $Email", "Add proxy addresses $proxyAddresses")) {
        Set-Mailbox -Identity $Email `
                    -EmailAddresses @{add = $proxyAddresses} `
                    -Confirm:$false
    }

    # Gives the permission to new owners.
    Add-FullAccessPermission -Email $Email -FullAccessEmails $FullAccessEmails
    Add-ReviewerPermission -Email $Email -ReviewerEmails $ReviewerEmails

    # Restore the old mailbox to the new one.
    Restore-OldToNewMailbox -Email $Email
}


#┌───────────────────┐
#│    MAIN PROCESS   │
#└───────────────────┘
# Delete the user from the AD
if($DeleteAD) { Delete-ADUserFromEmail -Email $Email }

if($Archive) {
    # Wait for the user deletion in Exchange
    if($PSCmdlet.ShouldProcess("Mailbox $Email", "Waiting for the deletion in Exchange")) {
        Wait-UserDeletionExchange -Email $Email -MaxWaitMinutes $MaxWaitMinutes
    }
    
    # Create a new shared mailbox and restore the deleted singular mailbox to that shared mailbox
    CreateAndRestore-OldToNewMailbox -Email $Email `
                                     -ProxyFilter $ProxyFilter `
                                     -FullAccessEmails $FullAccessEmails `
                                     -ReviewerEmails $ReviewerEmails `
}

if($RedirectEmail) {
    # For external redirection
    if($RedirectExternal) {
        if($PSCmdlet.ShouldProcess("Mailbox $Email", "Set external forwarding to $RedirectEmail")) {
            Set-Mailbox -Identity $Email `
                        -ForwardingSMTPAddress $RedirectEmail `
                        -DeliverToMailboxAndForward $DeliverToMailboxAndForward
        }
    }
    # For same tenant redirection
    else {
        if($PSCmdlet.ShouldProcess("Mailbox $Email", "Set internal forwarding to $RedirectEmail")) {
            Set-Mailbox -Identity $Email `
                        -ForwardingAddress $RedirectEmail `
                        -DeliverToMailboxAndForward $DeliverToMailboxAndForward
        }
    }
}

Disconnect-ExchangeOnline -Confirm:$false
