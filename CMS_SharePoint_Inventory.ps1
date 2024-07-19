<#
.SYNOPSIS
This script processes user information from a CSV file, checks SharePoint site memberships for each user, and exports the results to a new CSV file.

.DESCRIPTION
The script ensures that the Microsoft.SharePoint.PowerShell module is installed and imported, reads a list of users from a CSV file, connects to a SharePoint Online tenant, retrieves all SharePoint sites, checks each user's membership across these sites, and exports the details to a new CSV file.

.NOTES
The script requires an account with SharePoint Online Administrator permissions for authentication.

.AUTHOR
SubjectData

.EXAMPLE
.\CMS_SharePoint_Inventory.ps1
This will run the script in the current directory, processing the 'Users.csv' file and generating 'CMSSPSites.csv' with SharePoint site details for each user in the list.
#>

$MicrosoftOnline = "Microsoft.Online.Sharepoint.PowerShell"

# Check if the module is already installed
if (-not(Get-Module -Name $MicrosoftOnline)) {
    # Install the module
    Install-Module -Name $MicrosoftOnline -Force
}

Import-Module $MicrosoftOnline -Force

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$XLloc = "$myDir\"
$ReportsPath = "$myDir\"

$results = @()

try
{
    $UsersList = import-csv ($XLloc+"Users.csv").ToString()
}
catch
{
    Write-Host "No CSV file to read" -BackgroundColor Black -ForegroundColor Red
    exit
}


########################################
###   Connect To SharePoint Tenant   ###
########################################

$tenantUrl = "https://m365x72261057-admin.sharepoint.com"

Connect-SPOService -url $tenantUrl


#Fetch All SharePoint sites

$AllSites = Get-SPOSite -Limit All

if($UsersList.Count -gt 0)
{
    foreach ($CMSUser in $UsersList)
    {   
        try
        {
            Write-Host "Processing User $($CMSUser.'UserEmail')"

            foreach($currentSite in $AllSites)
            {
                Write-Host "Processing Team $($currentSite.Url)"

                $queryUser = Get-SPOUser -Site $currentSite.Url | ?{$_.LoginName -eq $CMSUser.'UserEmail'}

                #Get-SPOUser -Site "" | ?{$_.LoginName -eq "" -and $_.Groups.Count -gt 0}
                        
                if($queryUser.Count -gt 0)
                {            
                    $spUsers = Get-SPOUser -Site $currentSite.Url
                        
                    $details = @{            
                    
                        SharePointSite       = $currentSite.Url

                        SiteUsersEmail       = $spUsers.LoginName -join ";"

                        SiteUsersName        = $spUsers.DisplayName -join ","

                        CMSUser              = $CMSUser.'UserEmail'                        
                             
                    }

                    $results += New-Object PSObject -Property $details   
                }
            }
        }
        catch
        {
            Write-Host "Exception occured " $currentSite.DisplayName -BackgroundColor Black -ForegroundColor Red
            continue
        }
    }
}

$results | export-csv -Path "$($XLloc)\CMSSPSites.csv" -NoTypeInformation


