# Import active directory module for running AD cmdlets
Import-Module activedirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv C:\temp\userlist.csv

#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$Username 	= $User.SamName
	$Firstname 	= $User.First
	$Lastname 	= $User.Last
	$Description = $User.Description
	$UPN        = "$($User.First).$($User.Last)@contoso.com"
	$OU 		= "OU=Users,DC=contoso,DC=com"
    $email      = $User.Email
    $city       = $User.City
    $country    = $User.Country
    $office     = $User.Office
    $jobtitle   = $User.JobTitle
    $company    = $User.Company
    $displayname = $User.DisplayName
	$password    = "Password" | ConvertTo-SecureString -AsPlainText -Force


	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account
		
        #Account will be created in the OU provided by the $OU variable read from the CSV file
		New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName $UPN `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
			-ChangePasswordAtLogon $false `
			-Description $User.Description `
            -DisplayName $displayname `
            -Path $OU `
            -City $city `
            -Company $company `
			-Office $office `
            -EmailAddress $email `
            -Title $jobtitle `
            -AccountPassword $password `
            
	}
}