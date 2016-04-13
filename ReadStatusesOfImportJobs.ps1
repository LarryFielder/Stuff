<# 

# User profile bulk API - Read all import jobs from the tenant # 
# 13th of Apr, 2016 - Release v1.0
# Author(s)
Nano Nano

# Notes
Support for this script requires that you have installed update SharePoint Online SDK/CSOM redistributable
with minimum version of 4622.1208 to the computer where the script is executed. 
Download redistributable package from https://www.microsoft.com/en-us/download/details.aspx?id=35585
#>


# Load assemblies to PowerShell session - Will try to resolve tenant dll status automatically, if possible
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
$a = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Online.SharePoint.Client.Tenant")
if( !$a ){
    # Let's try to load that from default location.
    $defaultPath = "C:\Program Files\SharePoint Client Components\16.0\Assemblies\Microsoft.Online.SharePoint.Client.Tenant.dll"
    $a = [System.Reflection.Assembly]::LoadFile($defaultPath)
}


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
