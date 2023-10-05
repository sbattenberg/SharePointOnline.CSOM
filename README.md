# SharePointOnline.CSOM

This module allows the usage of the SharePoint Online Client-Side Object Model (CSOM) libraries at Azure Runbooks. This module uses the libraries from the NuGet package "[Microsoft.SharePointOnline.CSOM](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/)".

## Functions

**Load-SPOnlineCSOMAssemblies**

Loads the common CSOM libraries to the PowerShell session:
* Microsoft.SharePoint.Client.dll
* Microsoft.SharePoint.Client.Publishing.dll
* Microsoft.SharePoint.Client.Runtime.dll
* Microsoft.SharePoint.Client.Search.dll

```
Load-SPOnlineCSOMAssemblies
```

**Load-SPOnlineCSOMAssembly**

Loads a specific CSOM library to the PowerShell session. Available libraries:
* Microsoft.Office.Client.Policy.dll
* Microsoft.Office.Client.TranslationServices.dll
* Microsoft.Office.SharePoint.Tools.dll
* Microsoft.Online.SharePoint.Client.Tenant.dll
* Microsoft.ProjectServer.Client.dll
* Microsoft.SharePoint.Client.dll
* Microsoft.SharePoint.Client.DocumentManagement.dll
* Microsoft.SharePoint.Client.Publishing.dll
* Microsoft.SharePoint.Client.Runtime.dll
* Microsoft.SharePoint.Client.Search.Applications.dll
* Microsoft.SharePoint.Client.Search.dll
* Microsoft.SharePoint.Client.Taxonomy.dll
* Microsoft.SharePoint.Client.UserProfiles.dll
* Microsoft.SharePoint.Client.WorkflowServices.dll
```
Load-SPOnlineCSOMAssembly "Microsoft.SharePoint.Client.Taxonomy.dll"
```

## Example
```
Load-SPOnlineCSOMAssemblies

$psCredential = Get-AutomationPSCredential -Name 'O365Credential'
$username = $psCredential.UserName
$securePassword = $psCredential.Password

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
$siteURL = "https://mySharePoint.sharepoint.com"
$cContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
$cContext.AuthenticationMode = [Microsoft.SharePoint.Client.ClientAuthenticationMode]::Default
$cContext.Credentials = $credentials

$cweb = $cContext.Web
$csite = $cContext.Site
$cContext.Load($cweb)
$cContext.Load($csite)
$cContext.ExecuteQuery()
$cweb.Title
```
Create a credential asset for the Azure Automation account with the name 'O365Credential' and the SharePoint Online credentials.

## Links

- [PowerShell Gallery Download](https://www.powershellgallery.com/packages/SharePointOnline.CSOM)
- [Repository](https://github.com/sbattenberg/SharePointOnline.CSOM)
- [License Info](https://github.com/sbattenberg/SharePointOnline.CSOM/blob/main/LICENSE)




----------------
### Included third-party software/licensing:
Microsoft.SharePointOnline.CSOM (https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)
 - © Microsoft Corporation. All rights reserved (https://go.microsoft.com/fwlink/?LinkId=280198)

----------------
&copy; [sbattenberg](https://github.com/sbattenberg)
