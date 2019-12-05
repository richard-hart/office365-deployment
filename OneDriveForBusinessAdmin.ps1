<#
.SYNOPSIS
OneDrive for Business Management Tool

.DESCRIPTION
Configure OneDrive for Business permissions and remove folders across all OD4B
sites.

.PARAMETER BlockAccess
Block or unblock access to OneDrive for Business sites.

.PARAMETER Confirm
Use this switch parameter to confirm file or folder deletions.

.PARAMETER Credential
Used to store a credential object for connecting to SharePoint Online.

.PARAMETER DeleteFilePattern
Delete files matching pattern in a given OD4B site.  For example, deleteing all
MP4 files.  Must use Confirm switch to delete files; otherwise, log only. Use
a regular expression to ensure proper deletion.

.PARAMETER FilesModifiedOnThisDate
Use this parameter to specify the modified-on date for file version restores.
For example, if you have to restore the version prior to the last modification
5 dates ago, you could specifiy ((Get-Date).AddDays(-5)).

.PARAMETER FolderSize
Use this parameter to return the folder size for the specified users. Folder size
does NOT include version history storage (only final version).  Output in the 
"Note" column is in megabytes (MB).

.PARAMETER FolderToAdd
String parameter with the name of a folder to add to all matching OneDrive for
Business sites.  You must have permissions to create folders in the target
users' OneDrive for Business sites.  You can use the GrantPermissions parameter
to grant necessary permissions.

.PARAMETER FolderToDelete
String parameter witih the name of a folder to delete from all matching OneDrive
for Business sites.  You must have permissions to create folders in the target
users' OneDrive for Business sites.  You can use the GrantPermissions parameter
to grant necessary permissions. You'll need to use -Confirm if you actually want
to delete the folder.

.PARAMETER GrantPermissions
Grant permissions to a user account.  User is made a secondary site collection
administrator.  Useful for delegating permissions for eDiscovery.

.PARAMETER GrantPermissionsTo
Grant permissions to the user specified in this parameter.  If no user is
specified, use the value stored in the Credential.

.PARAMETER HoldStatus
Check for In-Place Holds applied to OD4B sites (beta).

.PARAMETER Identity
Specify an individual user (by UPN or email address) for OneDrive admin
operations.  If no identity is specified, run against all enumerated OneDrive
for Business sites.

.PARAMETER InputFile
List of UPNs or Email addresses for OneDrive for Business sites.  If no file
is specified, run against all enumerated OneDrive for Business sites.

.PARAMETER ListOneDriveSites
Generate a list of all OD4B sites.

.PARAMETER ListOneDriveSitesOutput
Output file for listing of OD4B sites.

.PARAMETER Logfile
Logfile for operations.

.PARAMETER RestoreVerions
Restore OneDrive for Business files in a site from saved versions.

.PARAMETER RevokePermissions
Revoke permissions for OneDrive for Business sites.

.PARAMETER RevokePermissionsFor
Revoke permissions of the named user.  If no user is specified, uses the value
stored in the credential.

.PARAMETER Tenant
SharePoint online tenant name (e.g., contoso or contoso.onmicrosoft.com)

.PARAMETER VersionsToGoBack
Specify number of versions of files to restore. By default, go back one version.

.EXAMPLE
.\OneDriveForBusinessAdmin.ps1 -Credential (Get-Credential) -GrantPermissions -GrantPermissionsTo aylakol@contoso.com

Grant permissions to aylakol@contoso.com to all OneDrive for Business sites.

.EXAMPLE
.\OneDriveForBusinessAdmin.ps1 -Credential (Get-Credential) -GrantPermissions -FolderToAdd "Sales Orders" -FolderToDelete "Shared with Everyone"

Grant permissions to the user specified in -Credential, add the folder "Sales
Orders" to the "Documents" document library, and remove the "Shared with
Everyone" folder from the "Documents" document library.

.EXAMPLE
.\OneDriveForBusinessAdmin.ps1 -Credential (Get-Credential) -GrantPermissions -DeleteFilePattern "\.mp4$" -Confirm

Grant permissions to the user specified in *Get-Credential) and delete files with extension matching .mp4.

.EXAMPLE
.\OneDriveForBusinessAdmin.ps1 -Credential (Get-Credential) -Tenant ems340903 -ListOneDriveSites

Generate a list of all OneDrive for Business sites and output to default location.

.LINK
https://aka.ms/OneDriveAdmin

.LINK
https://undocumented-features.com/2017/08/25/onedrive-for-business-admin-tool/

.LINK
https://undocumented-features.com/2017/10/16/recovering-from-crypto-or-ransomware-attacks-with-the-onedrive-for-business-admin-tool/

.NOTES
2019-08-01	Updates/clarity in messages based on user feedback, text alignment updates.
2019-05-15 	Added HoldStatus check (beta).
2018-10-04  Added ListOneDriveSites parameter.
		    Fixed Identity parameter bug when submitting multiple users via command-line interface.
2018-09-15	Added FolderSize parameter.
2018-01-10	Added DeleteFilePattern parameter. Updated LogFile parameter to always generate.
2017-10-16  Added Identity and RestoreVersions parameters
2017-09-05	Added TrimEnd("/") to BlockAccess function
2017-08-24	Initial release.
#>

Param (
	[ValidateSet('Block','Unblock')]
	[string]$BlockAccess,
	
	[Parameter(Mandatory = $false, HelpMessage = 'Confirm removal of the folders')]
	[switch]$Confirm,
	
	[Parameter(Mandatory = $true)]
	[System.Management.Automation.PSCredential]$Credential,
	
	[Parameter(Mandatory = $false)]
	[regex]$DeleteFilePattern,
	
	[Parameter(Mandatory = $false)]
	[datetime]$FilesModifiedOnThisDate,
	
	[Parameter(Mandatory = $false)]
	[switch]$FolderSize,
	
	[Parameter(Mandatory = $false, HelpMessage = 'Folder to add')]
	[string]$FolderToAdd,
	
	[Parameter(Mandatory = $false, HelpMessage = 'Folder to delete')]
	[string]$FolderToDelete,
	
	[Parameter(mandatory = $false)]
	[Switch]$GrantPermissions,
	
	[Parameter(Mandatory = $false)]
	[string]$GrantPermissionsTo,
	
	[switch]$HoldStatus,
	
	[array]$Identity,
	
	[String]$InputFile,
	
	[switch]$ListOneDriveSites,
	
	[string]$ListOneDriveSitesOutput = (Get-Date -Format yyyy-MM-dd) + "_OneDriveForBusinessSiteList.csv",
	
	[Parameter(mandatory = $false)]
	[String]$LogFile = (Get-Date -Format yyyy-MM-dd) + "_OneDriveForBusinessAdminLog.csv",
	
	[Parameter(Mandatory = $false)]
	[ValidateSet('ByDate', 'ByNumberOfVersionsToGoBack')]
	[String]$RestoreVersions,
	
	[Parameter(mandatory = $false)]
	[Switch]$RevokePermissions,
	
	[Parameter(Mandatory = $false)]
	[string]$RevokePermissionsFor,
	
	[Parameter(mandatory = $true)]
	[String]$Tenant,
	
	[Parameter(Mandatory = $false)]
	[int]$VersionsToGoBack
	
) # End Parameters
$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

If ($HoldStatus)
{
	If (!(Get-Command Get-CaseHoldPolicy -ea silently continue))
	{
		try
		{
			$ComplianceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Credential -Authentication Basic -AllowRedirection
			Import-PSSession $ComplianceSession -ea silentlycontinue -wa silentlycontinue
		}
		catch { }
		
		If (!(Get-Command Get-EDiscoveryCaseAdmin -ea silentlycontinue))
		{
			Write-Warning -Message "Current user is not a member of eDiscovery Case Admins. If holds are applied to OD4B sites as part of an eDiscovery case, the case identity may not be fully resolved, though the GUID will be visible."
		}
	}
}

function LogWrite($Function,$User,$OD4BSite,$FolderName,$Note,$VersionLabel)
{
	$LogData = """" + $Function + """" + "," + """" + $User + """"+ "," + """" + $OD4BSite + """" + "," + """" + $FolderName + """" + "," + """" + $Note + """" + "," + """" + $VersionLabel + """"
	Add-Content -Path $LogFile -Value $LogData
} # End Function LogWrite

function ListOneDriveSites($OD4BSite, $User)
{
	$OD4BSite = $OD4BSite.TrimEnd("/")
	$Data = """" + $($User) + """" + "," + """" + $OD4BSite + """"
	LogWrite -Function "ListOneDriveSites" -User $User -OD4BSite $OD4BSite  
	Add-Content -Path $ListOneDriveSitesOutput -Value $Data
} # End Function ListOneDriveSites

function BlockOneDriveAccess($OD4BSite, $AccessState)
{
	$OD4BSite = $OD4BSite.TrimEnd("/")
	Switch ($AccessState)
	{
		Block
		{
			Write-Host "AccessState Should be set to NoAccess."
			Get-SPOSite -Identity $OD4BSite | Set-SPOSite -LockState NoAccess
			If ($LogFile) { LogWrite -Function "$($MyInvocation.MyCommand)-$AccessState.ToString()" -OD4BSite $OD4BSite }
			$State = ((Get-SPOSite -Identity $OD4BSite).LockState)
			Write-Host "Access state is: " $State
			Write-Host "-----------------------------------------------------------------------"
		}
		
		Unblock
		{
			Write-Host "AccessState Should be set to Unlock."
			Get-SPOSite -Identity $OD4BSite | Set-SPOSite -LockState Unlock
			If ($LogFile) { LogWrite -Function "$($MyInvocation.MyCommand)-$AccessState.ToString()" -OD4BSite $OD4BSite }
			$State = ((Get-SPOSite -Identity $OD4BSite).LockState)
			Write-Host "Access state is: " $State
			Write-Host "-----------------------------------------------------------------------"
		}
	}
} # End Function BlockOneDriveAccess

Function GrantPermissions($User, $OD4BSite)
{
	Write-Host -ForegroundColor Green "     Granting permissions on $OD4BSite to $User"
	Set-SPOUser -Site $OD4BSite -LoginName $User -IsSiteCollectionAdmin $true | Out-Null
	If ($LogFile) { LogWrite -Function $MyInvocation.MyCommand -User "$($User)" -OD4BSite "$($OD4BSite)" }
} # End Function GrantPermissions

Function RevokePermissions($User, $OD4BSite)
{
	Write-Host -ForegroundColor Green "     Revoking permissions on $($OD4BSite) for $($User)"
	Get-SPOUser -Site $OD4BSite.TrimEnd("/")
	If ($LogFile) { LogWrite -Function $MyInvocation.MyCommand -User "$($User)" -OD4BSite "$($OD4BSite)" }
} # End Function RevokePermissions

Function AddFolder($FolderName,$OD4BSite)
{
	Write-Host "Adding Folder Documents/$($FolderName)"
	$Folder = $personalWeb.Folders.Add("Documents" + "/" + $FolderName)
	$ClientContextSource.Load($personalWeb)
	$ClientContextSource.Load($personalWeb.Folders)
	$ClientContextSource.ExecuteQuery()
	If ($LogFile) { LogWrite -Function $MyInvocation.MyCommand -FolderName "$($FolderName)" -OD4BSite "$($OD4BSite)" }
	
} # End Function AddFolder

Function RemoveFolder($FolderName,$OD4BSite,$User)
{
# Delete Specified Folder
foreach ($toBeDeleted in $allFolders)
	{
		#Write-Host "Examining Folder" $toBeDeleted.Name
		if ($toBeDeleted.Name -eq $FolderName)
		{
			#Write-Host -Fore Green "     $($FolderName) present in $OD4BPath"
			If ($Confirm)
			{
				Write-Host -ForegroundColor Cyan "     ** Confirm enabled. Deleting Folder " $toBeDeleted.Name
				$toBeDeleted.DeleteObject()
				$personalWeb.Update()
				$ClientContextSource.ExecuteQuery()
				If ($LogFile)
				{
					LogWrite -Function $MyInvocation.MyCommand -FolderName "$($FolderName)" -OD4BSite "$($OD4BSite)"
				}
			}
			Else
			{
				Write-Host -ForegroundColor Yellow "No action taken on $($toBeDeleted). Use -Confirm parameter to delete."
			}
		}
	}
} # End Function RemoveFolder

function DeleteFilePattern($Pattern, $OD4BSite)
{
	If ($Confirm)
	{
		$global:FilesToDelete = @()
		foreach ($obj in $allItems)
		{
			If ($obj["FileRef"] -match $Pattern)
			{
				#write-host "File $($obj["FileRef"]) matches. Deleting."
				#$obj.DeleteObject())
				$FilesToDelete += $obj
				#$ClientContextSource.ExecuteQuery()
				If ($LogFile) { LogWrite -Function "DeleteFilePattern" -OD4BSite $OD4BSite -Note "File $($obj["FileRef"]) deleted." }
			}
		}
		$FilesToDelete.DeleteObject()
		$ClientContextSource.ExecuteQuery()
	}
	Else
	{
		foreach ($obj in $allItems)
		{
			If ($obj["FileRef"] -match $Pattern)
			{
				write-host "File $($obj["FileRef"]) matches. Specify Confirm switch to delete.";
				If ($LogFile) { LogWrite -Function "DeleteFilePattern-LogOnly" -OD4BSite $OD4BSite -Note "File $($obj["FileRef"]) would have been deleted." }
			}
		}
	}
}

function HoldStatus($OD4BSite, $User)
{
	# Explicit holds
	$LibraryStatus = $personalWeb.AllProperties
	$ClientContextSource.Load($LibraryStatus)
	$ClientContextSource.ExecuteQuery()
	[array]$holdstemp = $LibraryStatus.FieldValues.allwebholds -split ";"
	$holds = @()
	foreach ($obj in $holdstemp)
	{
		$o = $obj.Split(",")[0]
		$holds += $o
	}
	[pscustomobject]$HoldPolicies = @()
	foreach ($policy in $holds)
	{
		If (!($policy -eq ""))
		{
			$HoldPolicies = Get-RetentionCompliancePolicy $policy | select `
																		   @{ N = "Hold"; E = { $_.Name } },
																		   @{ N = "Policy Guid"; E = { $policy } }
			If (!($HoldPolicies.Hold -eq "")) { $HoldString = "$($HoldPolicies.Hold) - $($HoldPolicies.'Policy Guid')" }
			else { $HoldString = "<unable to resolve> - $($HoldPolicies.'Policy Guid')"}
			Write-Host "$($OD4BSite) has the following holds applied to it: $HoldString"
			LogWrite -Function "HoldStatus" -User $User -OD4BSite $OD4BSite -Note $HoldString
		}
	}
}

function FolderSize($OD4BSite,$User)
{
	If ($docList.BaseType -eq "DocumentLibrary")
	{
		$BatchSize = 1000
		$TotalLibrarySize = 0
		$FileCount = 0
		$docSize = 0
		$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
		$camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit Paged='True'>$BatchSize</RowLimit></View>";
		$allLibraryItems = $doclist.GetItems($camlQuery)
		$ClientContextSource.Load($allLibraryItems)
		$ClientContextSource.ExecuteQuery()
		
		foreach ($item in $allLibraryItems)
		{
			$FileCount++
			if ($item.FileSystemObjectType -eq "File")
			{
				$file = $item.File
				$fItem = $file.ListItemAllFields
				$ClientContextSource.Load($file)
				$ClientContextSource.Load($fItem)
				
				$ClientContextSource.ExecuteQuery()
				$docSize = $fItem["File_x0020_Size"]
				[int]$TotalLibrarySize += $docSize
				
				# To output each file individually to screen for troubleshooting, uncomment
				#Write-Host "$($docList.Title), $($file.ServerRelativeUrl), $($fItem["FileLeafRef"].Split('.')[0]), $($docSize) KB"
			}
		}
		$TotalLibrarySize = $TotalLibrarySize/1048576 
		Write-Host "$($User) folder total size is $($TotalLibrarySize) MB in $($FileCount) files."
		LogWrite -Function FolderSize -User $($User) -OD4BSite $OD4BSite -Note $TotalLibrarySize
	}
	
	##
} # End Function FolderSize

### Script component prerequisites 
Write-Host -Fore Yellow "Locating SharePoint Server Client Components installation..."
If (Test-Path 'c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll')
{
	Write-Host -ForegroundColor Green "Found SharePoint Server Client Components installation."
	Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
	Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
	Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
	Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"
}
ElseIf ($filename = (Get-ChildItem 'C:\Program Files' -Recurse -ea silentlycontinue | where { $_.name -eq 'Microsoft.SharePoint.Client.DocumentManagement.dll' })[0])
{
	$Directory = ($filename.DirectoryName)[0]
	Write-Host -ForegroundColor Green "Found SharePoint Server Client Components at $Directory."
	Add-Type -Path "$Directory\Microsoft.SharePoint.Client.dll"
	Add-Type -Path "$Directory\Microsoft.SharePoint.Client.Runtime.dll"
	Add-Type -Path "$Directory\Microsoft.SharePoint.Client.Taxonomy.dll"
	Add-Type -Path "$Directory\Microsoft.SharePoint.Client.UserProfiles.dll"
}

ElseIf (!(Test-Path 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'))
{
	Write-Host -ForegroundColor Yellow "This script requires the SharePoint Server Client Components. Attempting to download and install."
	wget 'https://download.microsoft.com/download/E/1/9/E1987F6C-4D0A-4918-AEFE-12105B59FF6A/sharepointclientcomponents_15-4711-1001_x64_en-us.msi' -OutFile ./SharePointClientComponents_15.msi
	wget 'https://download.microsoft.com/download/F/A/3/FA3B7088-624A-49A6-826E-5EF2CE9095DA/sharepointclientcomponents_16-4351-1000_x64_en-us.msi' -OutFile ./SharePointClientComponents_16.msi
	msiexec /i SharePointClientComponents_15.msi /qb
	msiexec /i SharePointClientComponents_16.msi /qb
	Sleep 60
	If (Test-Path 'c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll')
	{
		Write-Host -ForegroundColor Green "Found SharePoint Server Client Components."
		Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
		Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
		Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
		Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"
	}
	Else
	{
		Write-Host -NoNewLine -ForegroundColor Red "Please download the SharePoint Server Client Components from "
		Write-Host -NoNewLine -ForegroundColor Yellow "https://download.microsoft.com/download/F/A/3/FA3B7088-624A-49A6-826E-5EF2CE9095DA/sharepointclientcomponents_16-4351-1000_x64_en-us.msi "
		Write-Host -ForegroundColor Red "and try again."
		Break
	}
}

If (!(Get-Module -ListAvailable "*online.sharepoint*"))
{
	Write-Host -ForegroundColor Yellow "This script requires the SharePoint Online Management Shell.  Attempting to download and install."
	wget 'https://download.microsoft.com/download/0/2/E/02E7E5BA-2190-44A8-B407-BC73CA0D6B87/SharePointOnlineManagementShell_6802-1200_x64_en-us.msi' -OutFile ./SharePointOnlineManagementShell.msi
	msiexec /i SharePointOnlineManagementShell.msi /qb
	Write-Host -ForegroundColor Yellow "Please close and reopen the Windows Azure PowerShell module and re-run this script."
}
### End script component prerequisites

If ($InputFile -and $Identity)
{
	Write-Host -ForegroundColor Red "Only InputFile or Identity may be specified on the command line (or neither to include all users)."
	Break
}

If ($LogFile)
{
	If (!(Test-Path $LogFile))
	{
		Write-Host -ForegroundColor Yellow "Log file not found. Creating $($LogFile)."
		# Params that can be passed to LogWrite: $Function, $User, $OD4BSite, $FolderName, $Note, $VersionLabel
		$LogFileHeader = """" + "Function" + """" + "," + """" + "User" + """" + "," + """" + "OD4BSite" + """" + "," + """" + "FolderName" + """" + "," + """" + "Note" + """" + "," + """" + "VersionLabel" + """"
		
		$LogFileHeader | Out-File $LogFile
	}
	Else
	{
		Write-Host -ForegroundColor Yellow "Existing log file found. Appending."
	}
}

If ($ListOneDriveSites)
{
	If (!(Test-Path $ListOneDriveSitesOutput))
	{
		Write-Host -ForegroundColor Yellow "One Drive Site List output file not found. Creating $($ListOneDriveSitesOutput)."
		$ListHeader = """" + "UserPrincipalName" + """" + "," + """" + "OD4BSite" + """"
		$ListHeader | Out-File $ListOneDriveSitesOutput
	}
}

## Variables
# Define URLs
If ($tenant -like "*.onmicrosoft.com") { $tenant = $tenant.split(".")[0] }
$MySiteURL = "https://$tenant-my.sharepoint.com"
$AdminURL = "https://$tenant-admin.sharepoint.com"

# Define Contexts
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($MySiteURL)

# Define Credentials
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.Password)
$Context.Credentials = $Creds

# Connect to SPO Service for granting permissions if necessary
Connect-SPOService -url $AdminURL -credential $Credential

# Get OD4B WebSite Users
$Users = $Context.Web.SiteUsers
$Context.Load($Users)
$Context.ExecuteQuery()
$peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Context)

# Set user list to $Context.WebSite.Users if no input file is found
If ($InputFile)
{
	$UserList = Import-Csv -Header UserPrincipalName
	$UserList = $UserList.UserPrincipalName -join "|"
	[array]$users = $Users.LoginName | ? { $_ -match $UserList }
}
If ($Identity)
{
	If ($Identity.Count -gt 1) { $UserList = $Identity -join "|" }
	else { $UserList = $Identity }
	#Write-Host "UserList is $($UserList)"
	[array]$Users = $Users.LoginName | ? { $_ -match $UserList }
}

# Define GrantPermissionsTo and RevokePermissionsFrom Users
If (!$GrantPermissionsTo) { $GrantPermissionsTo = $Credential.UserName }
If (!$RevokePermissionsFor) { $RevokePermissionsFor = $Credential.UserName }

foreach ($User in $Users)
{
	# select $User to load into $UserProfile
	If ($Identity)
	{
		$userProfile = $peopleManager.GetPropertiesFor($User)
	}
	
	Else
	{
		$userProfile = $peopleManager.GetPropertiesFor($user.LoginName)
	}
	
	# Update $userProfile var
	$Context.Load($userProfile)
	$Context.ExecuteQuery()
	
	# Check to see if the User has an email address and a OD4B site provisioned
	if ($userProfile.Email -ne $null -and $userProfile.UserProfileProperties.PersonalSpace -ne "")
	{
		$i++
		$OD4BPath = $MySiteURL + $userProfile.UserProfileProperties.PersonalSpace
		Write-Host -ForegroundColor Green "Processing $($OD4BPath) for $($userProfile.UserProfileProperties.PreferredName)..." # $userProfile.UserProfileProperties.PersonalSpace
		
		# If the $GrantPermissions switch is present, execute the GrantPermissions function
		If ($GrantPermissions)
		{
			GrantPermissions -User $GrantPermissionsTo -OD4BSite $OD4BPath
		}
		
		If ($RevokePermissions)
		{
			RevokePermissions -User $RevokePermissionsFor -OD4BSite $OD4BPath
		}
		
		If ($BlockAccess)
		{
			BlockOneDriveAccess -OD4BSite $OD4BPath -AccessState $BlockAccess	
		}
		
		# If OD4BPath is present, enumerate folders for various operations
		if (($HoldStatus -or $FolderToAdd -or $FolderToDelete -or $DeleteFilePattern -or $FolderSize) -and $OD4BPath)
		{
			$ClientContextSource = New-Object Microsoft.SharePoint.Client.ClientContext($OD4BPath);
			
			$ClientContextSource.Credentials = $Creds
			$ClientContextSource.ExecuteQuery()
			
			$personalWeb = $ClientContextSource.Web
			$ClientContextSource.Load($personalWeb)
			$ClientContextSource.ExecuteQuery()
			
			$docList = $personalWeb.Lists.GetByTitle("Documents")
			$ClientContextSource.Load($docList)
			$ClientContextSource.Load($personalWeb.Folders)
			$ClientContextSource.ExecuteQuery()
			
			$allFolders = $docList.RootFolder.Folders
			$ClientContextSource.Load($allFolders)
			$ClientContextSource.ExecuteQuery()
			
			# Get Hold status
			If ($HoldStatus)
			{
				HoldStatus -OD4BSite $OD4BPath -User $($userProfile.Email)
			}
			
			# Get Folder Size
			If ($FolderSize)
			{
				FolderSize -OD4BSite $OD4BPath -User $($userProfile.Email)
			}
			
			# Remove Folder
			If ($FolderToDelete)
			{
				RemoveFolder -FolderName $FolderToDelete
			}
			
			# Add folder
			If ($FolderToAdd)
			{
				AddFolder -FolderName $FolderToAdd
			}
			
			# Delete files matching pattern
			If ($DeleteFilePattern)
			{
				$Library = $ClientContextSource.Web.Lists.GetByTitle("Documents")
				$ClientContextSource.Load($Library)
				$ClientContextSource.ExecuteQuery()
				$CamlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
				$CamlQuery.ViewXml = "<View Scope='RecursiveAll' />"
				$AllItems = $Library.GetItems($camlQuery)
				$ClientContextSource.Load($AllItems)
				$ClientContextSource.ExecuteQuery()
				
				If ($Confirm) { DeleteFilePattern -Pattern $DeleteFilePattern -Confirm }
				Else { DeleteFilePattern -Pattern $DeleteFilePattern }
			} # End DeleteFilePattern
	
		} # End If FolderToAdd or FolderToDelete or DeleteFilePattern
		
		If ($RestoreVersions)
		{
			# Load items and versions
			$ClientContextSource = New-Object Microsoft.SharePoint.Client.ClientContext($OD4BPath);
			
			$ClientContextSource.Credentials = $Creds
			$ClientContextSource.ExecuteQuery()
			
			$personalWeb = $ClientContextSource.Web
			$ClientContextSource.Load($personalWeb)
			$ClientContextSource.ExecuteQuery()
			
			$docList = $personalWeb.Lists.GetByTitle("Documents")
			$ClientContextSource.Load($docList)
			$ClientContextSource.Load($personalWeb.Folders)
			$ClientContextSource.ExecuteQuery()
			
			$allFolders = $docList.RootFolder.Folders
			$ClientContextSource.Load($allFolders)
			$ClientContextSource.ExecuteQuery()
			
			$VersionQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
			$VersionQuery.ViewXml = "<View Scope = 'RecursiveAll' />"
			$ItemsToRestore = $docList.GetItems($VersionQuery)
			$ClientContextSource.Load($ItemsToRestore)
			$ClientContextSource.ExecuteQuery()
			
			foreach ($item in $ItemsToRestore)
			{
				$File = $ClientContextSource.Web.GetFileByServerRelativeUrl($item["FileRef"])
				$ClientContextSource.Load($File)
				$ClientContextSource.Load($File.Versions)
				#$ClientContextSource.Load($File.ServerRelativeUrl)
				###
				
				Try
				{
					$ClientContextSource.ExecuteQuery()
					Write-Host File versions $File.Versions.Count $File.ServerRelativeUrl
				}
				Catch
				{
					Continue;
				}
				If ($File.Versions.Count -eq 0)
				{
					Write-Host No vesions to restore.
					$LogData = New-Object PSCustomObject
					$LogData | Add-Member NoteProperty ServerRelativeUrl($File.ServerRelativeUrl)
					$LogData | Add-Member NoteProperty FileLeafRef($item["FileLeafRef"])
					$LogData | Add-Member NoteProperty Versions("No Versions Available")
					If ($LogFile)
					{
						LogWrite -Function "RestoreVersionsFailed" -User $User -Note $LogData.ServerRelativeUrl
					}
				}
				ElseIf ($File.TypedObject.ToString() -eq "Microsoft.SharePoint.Client.File")
				{
					#foreach ($Version in $File.Versions)
					#{
					#$File.Versions | FL
					
					Switch($RestoreVersions)
					{
						ByDate
						{
							If (!($FilesModifiedOnThisDate))
							{
								Write-Host -ForegroundColor Red "You must specify a value for FilesModifiedOnThisDate."
								Break
							}
							
							[array]$FileRestoreCandidates = $File.Versions | ? { $_.Created -le $FilesModifiedOnThisDate } | Sort -Property Created
							If ($FileRestoreCandidates.Count -lt 2)
							{
								Write-Host -ForegroundColor Red "No versions available matching date criteria for $($File.ServerRelativeUrl)."
								Continue
							}
							
							Else
							{
								# Select the most current version in the
								$VersionLabel = $FileRestoreCandidates[-2].VersionLabel
								Write-Host "Version to be restored: " $VersionLabel
							}
							
							$File.Versions.RestoreByLabel($VersionLabel)
							$ClientContextSource.ExecuteQuery()
							
							If ($LogFile)
							{
								LogWrite -Function "FileRestored" -User $User -Note $LogData.ServerRelativeUrl -VersionLabel $VersionLabel
							}
						}
						
						ByNumberOfVersionsToGoBack
						{
							if ($File.Versions[($File.Versions.Count - 1)].IsCurrentVersion)
							{
								Write-Host $File.Name $Version.Created $Version.Size $Version.VersionLabel $Version.IsCurrentVersion $File.Versions.Count
								$VersionLabel = $File.Versions[($File.Versions.Count - 2)].VersionLabel
								Write-Host "Version to be restored: " $VersionLabel
							}
							elseif ($VersionsToGoBack -and $VersionsToGoBack -le $File.Versions.Count)
							{
								Write-Host "Versions to go back specified as $($VersionsToGoBack)"
								$VersionLabel = $File.Versions[($File.Versions.Count - $VersionsToGoBack)].VersionLabel
								Write-Host "Version to be restored: " $VersionLabel
							}
							else
							{
								$VersionLabel = $File.Versions[($File.Versions.Count - 2)].VersionLabel
								Write-Host "Version to be restored: " $VersionLabel
							}
							
							$File.Versions.RestoreByLabel($VersionLabel)
							$ClientContextSource.ExecuteQuery()
							
							If ($LogFile)
							{
								LogWrite -Function "FileRestored" -User $User -Note $LogData.ServerRelativeUrl -VersionLabel $VersionLabel
							}
						} # End Select "ByNumberOfversionsToGoBack"
					} # End Select
					#} #End Foreach $Version
				} # End ElseIf
			} # End Foreach $Item in $ItemsToRestore
		} # End RestoreVersions
		
		If ($ListOneDriveSites)
		{
			ListOneDriveSites -User $($userProfile.Email) -OD4BSite $OD4BPath
		}
	}
}