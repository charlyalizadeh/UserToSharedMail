# UserToSharedMail

## Description

Transfere the personal mailbox of a deleted user to a shared mailbox.
The default behavior of Microsoft is to delete the shared mailbox linked to a user if the user is deleted. This scripts allows the shared mailbox to stay active even after the user is deleted.
NOTE: the user must have been deleted in the last 30 days, during this time he is in a softdeleted state. After that the user no longer exists.

To keep the shared mailbox you have to:

1. Request the addresses and filter only the one ending with your domain
2. Delete the user from Active Directory
3. Wait for the user to be deleted from Exchange
4. Create a new shared mailbox with the old mail
5. Add proxy addresses to the new shared mailbox
6. Assigns permissions to new owners/reviewer
7. Get the GUID from the old mailbox and the new mailbox (even though they have the same name you can get both, see the script)
8. Restore the old mailbox to the new one

## Setup to run as a user

If you run this script as user the only thing you need is to have the correct right on Microsft Entra ID (and Active Directory if you want the script to also delete the user from the AD).

| Component        | Scope | Role                   | Permissions                                                   |
|------------------|-------|------------------------|---------------------------------------------------------------|
| Active Directory | User  | Domain Administrator   | Delete a user from Active Directory                           |
| Entra ID         | User  | Exchange Administrator | Create shared mailbox and manage proxy addresses, permissions |


## Running

```powershell
.\UserToSharedMail -Email prenom.nom@domain.com `
                   -ProxyFilter "^(smtp:|SMTP:).*@domain\.(com|fr)" `
                   -FullAccessEmails a.b@domain.com, c.d@domain.com `
                   -ReviewerEmails e.f@domain.com, g.h@domain.com `
                   -MaxWaitMinutes 30 `
                   -DeleteAD `
```

You can use the `-WhatIf` parameter to see what the would do:

```powershell
.\UserToSharedMail -Email prenom.nom@domain.com `
                   -ProxyFilter "^(smtp:|SMTP:).*@domain\.(com|fr)" `
                   -FullAccessEmails a.b@domain.com, c.d@domain.com `
                   -ReviewerEmails e.f@domain.com, g.h@domain.com `
                   -MaxWaitMinutes 30 `
                   -DeleteAD `
                   -WhatIf
```
