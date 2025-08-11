# UserToSharedMail

## Description

Script to convert a singular mailbox to a shared one. The Exchange Admin Center offers a way
to do it but it deletes the shared mailbox if the original user is deleted.
So to keep the shared mailbox you have to:

1. (Optional) Delete the user from Active Directory
2. (Optional) Wait for the user to be deleted from Microsoft 365 and Exchange
3. Request the proxy addresses
4. Create a new shared mailbox with the old mail
5. Add proxy addresses to the new shared mailbox
6. Assigns permissions to new owners
7. Restore the old mailbox to the new one


## Setup

Steps:
1. Register the application in Entra ID and assign the correct API permissions
2. Create the certificates and add them to Entra ID
3. Create a service principal in **Exchange** linked to the service principal in Entra ID (the one in entra should be automatically created after step 1)
4. Assign the right management role


### 1. Register the application in Entra ID and assign the correct API permissions

For this app to work you need to setup it in Entra ID(Azure AD) and Exchange:
1. Go to [Microsoft Entra admin center](https://entra.microsoft.com/)
2. `App registrations`
3. `New registration`
4. Enter the desired application name (it should be UserToSharedMail for better clarity) then press `Register`

To give the right permission to our app go to the application page in Entra:
1. `API permissions`
3. `Add a permission` -> `APIs my organization uses` -> `Office 365 Exchange Online` -> `Application permissions` -> `Exchange.ManageAsApp` -> `Add permissions`

If you have the `Application Administrator` rights in Entra (or `Global Administrator`) you can grant the permission yourself by clicking `Grant admin consent for axeria-iard.fr`.
If not you have to ask an administrator for consent.


### 2. Create the certificates and add them to Entra ID

To create the app certificates you need to run some PowerShell code and upload it to Entra:
```powershell
$cert = New-SelfSignedCertificate -Subject "CN=UserToSharedMailCertificate" `
                                  -CertStoreLocation "Cert:\CurrentUser\My" `
                                  -KeyExportPolicy Exportable `
                                  -KeySpec Signature `
                                  -KeyLength 2048 `
                                  -NotAfter (Get-Date).AddYears(3)
Export-Certificate -Cert $cert `
                   -FilePath "$env:USERPROFILE\Documents\ExchangeOnlineAutomation.cer"
Export-PfxCertificate -Cert $cert `
                      -FilePath "$env:USERPROFILE\Documents\ExchangeOnlineAutomation.pfx" `
                      -Password (Read-Host -AsSecureString)
```

Then go to the app page in Entra ID:
1. `Certificates & secrets`
2. `Certificates`
3. `Upload certificate`
4. Select the `.cer` file in your `Documents`


### 3. Create a service principal in Exchange linked to the service principal in Entra ID

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com/)
2. `Enterprise apps`
3. Search for the application name (`UserToSharedMail`)
4. Run (as an `Exchange Administrator`):
```powershell
New-ServicePrincipal -AppId <Application ID> -ObjectId <Object ID> -DisplayName <Name>`
```


### 4. Assign the right management role

Run (as an `Exchange Administrator`):
```powershell
# Get-Mailbox, New-Mailbox, Set-Mailbox
New-ManagementRoleAssignment -App <Application ID> -Role "Mail Recipients"
# Add-MailboxPermission, Add-MailboxFolderPermission
New-ManagementRoleAssignment -App <Application ID> -Role "Mail Recipients Creation"
# New-MailboxRestoreRequest
New-ManagementRoleAssignment -App <Application ID> -Role "Mail Import Export"
```


## Running

### As a user

```powershell
Connect-ExchangeOnline
.\UserToSharedMail.ps1 -Email example@domain.fr `
                       -ProxyFilter "^(smtp:|SMTP:).*@domain\.(fr|com)" `
                       -FullAccessEmails redirection1@domain.fr,redirection2@domain.fr `
                       -ReviewerEmails rewiewer1@domain.fr,rewiewer2@domain.com `
                       -DeleteAD $False `
                       -Archive $True
Disconnect-ExchangeOnline -Confirm:$false
```

### As an application

```powershell
$certificate = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -like "CN=UserToSharedMailCertificate" } # or a different name depending of how you called your certificate
Connect-ExchangeOnline -AppId <Application ID> `
                       -CertificateThumbprint $certificate.Thumbprint `
                       -Organization "<domain>.onmicrosoft.com"
.\UserToSharedMail.ps1 -Email example@domain.fr `
                       -ProxyFilter "^(smtp:|SMTP:).*@domain\.(fr|com)" `
                       -FullAccessEmails redirection1@domain.fr,redirection2@domain.fr `
                       -ReviewerEmails rewiewer1@domain.fr,rewiewer2@domain.com `
                       -DeleteAD $False `
                       -Archive $True
Disconnect-ExchangeOnline -Confirm:$false
```
