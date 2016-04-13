

# http://dev.office.com/blogs/introducing-bulk-upa-custom-profile-properties-update-api
 
# the path here may need to change if you used e.g. C:\Lib..
Add-Type -Path "C:\Code\CheckO365TenantVersion\packages\Microsoft.SharePointOnline.CSOM.16.1.5026.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Code\CheckO365TenantVersion\packages\Microsoft.SharePointOnline.CSOM.16.1.5026.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Code\CheckO365TenantVersion\packages\Microsoft.SharePointOnline.CSOM.16.1.5026.1200\lib\net45\Microsoft.Online.SharePoint.Client.Tenant.dll"


# Get needed information from end user
$adminUrl = Read-Host -Prompt 'Enter the admin URL of your tenant'
$userName = Read-Host -Prompt 'Enter your user name'
$pwd = Read-Host -Prompt 'Enter your password' -AsSecureString

# Get instances to the Office 365 tenant using CSOM
$uri = New-Object System.Uri -ArgumentList $adminUrl
$context = New-Object Microsoft.SharePoint.Client.ClientContext($uri)

$context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $pwd)
$o365 = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($context)
$context.Load($o365)

$jobs = $o365.GetImportProfilePropertyJobs()
$context.Load($jobs)
$context.ExecuteQuery();

foreach ($item in $jobs)
{
    Write-Host "ID: " $item.JobId " - Request status: "  $item.State " - Error status: "  $item.Error
}
