<#                   
.SYNOPSIS
    Create New User accounts from sharepoint list
.NOTES
    Authors:    Jake Cross
    Email:      jake@cross.house
    Updated:    2-5-23
#>
#############
#  Modules  #
#############
Import-Module PnP.PowerShell
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Users.Actions
Import-Module Microsoft.Graph.Authentication
Import-Module ActiveDirectory
###############
#  Variables  #
###############
# Sharepoint info for forms
$sharepoint = "INSERT HOME SHAREPOINT SITE URL etc. org.sharepoint.com"
$listName = "Employment"
# ID of application in Azure
$ClientID = "INSERT CLIENT ID OF APP REGISTERED IN AZURE"
# Thumbprint of certificate for authentication installed from setup process
$thumbprint = "INSERT CERTIFICATE THUMBPRINT FOR AUTH"
$cert = Get-ChildItem Cert:\LocalMachine\My\$Thumbprint
# 365 and on prem info
$oupath = "DN OF DEFAULT OU FOR USERS"
$Tenant = '_____ORGNAME____.onmicrosoft.com'
$tenantid = "TENANT ID"
$defaultPW = "DEFAULT PW"
$company = "COMPANY NAME"
$Website = "DOMAIN.COM"
#############
# Functions #
#############
# For adding a user to an array of groups
function Add-MemberToGroups {
    param (
        # Azure or AD Group
        [String] $GroupType,
        # Array of Groups
        [Array] $Groups,
        # UserID
        [String] $AADUserId
    )
    # For AD Groups
    if ($GroupType -eq "AD") {
        foreach ($group in $groups) {
            Add-ADGroupMember -Identity $group -members $sam
        }
    }
    # For Azure Groups
    elseif ($grouptype -eq "Azure") {
        foreach ($group in $groups) {
            # Get 365 groups and filter by 
            $mggroup = Get-MgGroup -filter "DisplayName eq '$group'"
            # Subject to add to group
            New-MgGroupMember -GroupId $mggroup.Id -DirectoryObjectId $AADUserId
        }
    }
}
# For getting assignments of a particular type
function Get-AssignmentsByType {
    param (
        # Assignment type like AD group or Azure Group
        [String] $AssignmentType,
        # Array of ids from assignments sharepoint list
        [Array] $GroupIds
    )
    # initializing array
    $AssignmentItems = @()
    # loop through all group ids in groups
    foreach ($groupid in $groupids) {
        # get assignment item based on id
        $AssignmentItem = get-pnplistitem -list 'Assignments' -id $groupid
        # add assignment item based on id if assignment matches assignment type
        if ($AssignmentItem.FieldValues.AssignmentType -eq $assignmenttype) { $AssignmentItems += $AssignmentItem }
    }
    # return array of assignments
    return $AssignmentItems.FieldValues.Title
}
#########################
# Connect to SharePoint #
#########################
# Gather new hire form information
# logs into sharepoint with certificate and App registered in Azure
Connect-PnPOnline -url $sharepoint -ClientId $ClientID -Tenant $Tenant -Thumbprint $thumbprint
# Get all users in employment list
$listitems = Get-PnPListItem -List $listName
# Grab only the accounts with approval states set to yes and created states set to no
$useraccounts = $listitems | where-object { $_.FieldValues.IsApproved -eq 'true' -and $_.FieldValues.IsCreated -ne 'true' }
##############
# Run Script #
##############
# Connect to Azure AD
Connect-MgGraph -ClientId $ClientID -TenantId $Tenantid -Certificate $cert
ForEach ($UserAccount in $UserAccounts) {
    # clears values from previous loops
    $Website = $null; $UPN = $null; $sam = $null; $guest = $null;
    # Grabs the job title of the current new hire
    $JobTitleObject = Get-PnPListItem -List 'Job Titles' -id $useraccount.FieldValues.Job_x0020_Title.LookupID
    $JobTitlegroupids = $JobTitleObject.FieldValues.Assignments.LookupID
    # Grabs the department of the current new hire from the job title
    $DepartmentObject = Get-PnPListItem -List 'Departments' -id $JobTitleObject.FieldValues.Department.LookupID
    $Departmentgroupids = $DepartmentObject.FieldValues.Assignments.LookupID
    # Grabbing License info to determine what kind of account to make
    $OfficeLicensePackage = $JobTitleObject.FieldValues.OfficeLicensePackage
    $Guest = $OfficeLicensePackage.contains("Guest")
    # Grabs the location of the new hire
    $Location = Get-PnPListItem -List 'Locations' -id $useraccount.FieldValues.Location.LookupID
    # Sets first name to preferred name if present
    if ($null -eq $UserAccount.FieldValues.PreferredName) {
        $FirstName = $UserAccount.FieldValues.FirstName
    }
    else {
        $FirstName = $UserAccount.FieldValues.PreferredName
    }
    if ($null -eq $UserAccount.FieldValues.PreferredLastName) {
        $LastName = $UserAccount.FieldValues.LastName
    }
    else {
        $LastName = $UserAccount.FieldValues.PreferredLastName
    }
    $dn = $FirstName + ' ' + $LastName
    $FirstName = $FirstName.trim()
    $LastName = $LastName.trim()
    # Creates Active Directory Account, syncs to 365, and assigns liceneses.
    if ($Guest -ne $true) {
        # Creating UserName to be first initial Last Name jcross - will need to add error processing later
        $sam = (($FirstName).Substring(0, 1) + $LastName)
        $sam = $sam.tolower()
        # Adding Office Location
        $address = $Location.FieldValues.Address
        $city = $Location.FieldValues.City
        $state = $Location.FieldValues.State
        $zip = $Location.FieldValues.Zip
        $Office = $useraccount.FieldValues.Location.LookupValue
        # Setting job title and department
        $jobtitle = $useraccount.FieldValues.Job_x0020_Title.LookupValue
        $department = $JobTitleObject.FieldValues.Department.LookupValue
        # Sets Manager
        $ManagerEmail = $useraccount.FieldValues.ReportingTo.Email
        $Manager = Get-ADUser -Filter { UserPrincipalName -eq $ManagerEmail }
        # Sets email suffix based on website
        $emailsuffix = "@" + $website
        # Sets UserPrincipleName of account
        $UPN = $sam + $emailSuffix
        #############################
        # Create Cloud Only Account #
        #############################
        if ($JobTitleObject.FieldValues.OnPremAccount -ne $true) {
            # Creates password object in format command accepts
            $PasswordProfile = @{
                Password = $defaultPW
            }
            # Create Cloud only user
            New-MgUser -GivenName $FirstName -Surname $LastName -DisplayName $dn -UserPrincipalName $UPN -Department $department -CompanyName $company -JobTitle $jobtitle `
                -PasswordProfile $PasswordProfile -OfficeLocation $Office -City $city -State $state -PostalCode $zip -StreetAddress $address -Country "US" `
                -AccountEnabled:$true -MailNickname $sam
        }
        ##########################
        # Create On Prem Account #
        ##########################
        else {
            # Creates AD Account
            New-ADuser -SamAccountName $sam -GivenName $FirstName -Surname $LastName -DisplayName $dn -Name $dn -accountPassword (ConvertTo-SecureString $defaultPW -AsPlainText -Force) `
                -UserPrincipalName $UPN -EmailAddress ($sam + $emailsuffix) -Path $oupath -Department $department -Company $company -HomePage $website -Manager $Manager.Samaccountname -Title $jobtitle `
                -Description $jobtitle -Office $Office -City $city -State $state -PostalCode $zip -StreetAddress $address -Country "US" -Enabled $true -ChangePasswordAtLogon $true
            # Adds user to groups based on department
            $DepartmentADGroups = get-assignmentsbytype -groupids $Departmentgroupids -AssignmentType "AD Group"
            Add-MemberToGroups -groups $DepartmentADGroups -GroupType "AD"
            # Adds user to groups based on job title
            $JobTitleADGroups = get-assignmentsbytype -groupids $JobTitlegroupids -AssignmentType "AD Group"
            Add-MemberToGroups -groups $JobTitleADGroups -GroupType "AD"
            # Start Sync with AD and Azure AD
            start-adsyncsynccycle -policytype delta
            # Wait 2 minutes for sync to take effect
            Start-Sleep -s 120
        }
        #######################################
        # For Both Cloud and On-Prem Accounts #
        #######################################
        # Set usage location - need to do before license is assigned, change US to your usage location
        Update-MgUser -UserId $UPN -UsageLocation 'US' 
        # Set office license based off of job title, array will work but only one primary license type can be assigned
        $LicensesToAssign = @()
        $LicensesToAssign = switch ($OfficeLicensePackage) {
            'Business Standard' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'Business Basic' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'F3' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'F1' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'Visio Web' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'Visio Desktop' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'Project' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
            'PowerBI' { @{skuid = 'XXXXXXXXXXXXXXXXXX' } }
        }
        # Assign License to user
        Set-MgUserLicense -UserId $UPN -AddLicenses $LicensesToAssign -RemoveLicenses @()
        # Assign Azure Groups to user based on department and job title
        $AADUser = get-mguser -UserId $UPN
        $JobTitleAzureGroups = get-assignmentsbytype -groupids $JobTitlegroupids -AssignmentType "Azure Group"
        $DepartmentAzureGroups = get-assignmentsbytype -groupids $Departmentgroupids -AssignmentType "Azure Group"
        Add-MemberToGroups -GroupType "Azure" -Groups $JobTitleAzureGroups -AADUserId $AADUser.Id
        Add-MemberToGroups -GroupType "Azure" -Groups $DepartmentAzureGroups -AADUserId $AADUser.Id
    }
    # Creates guest account in Azure AD.
    elseif ($Guest -eq $true) {
        New-MgInvitation -InvitedUserDisplayName $dn -InvitedUserEmailAddress $UserAccount.FieldValues.PersonalEmail -InviteRedirectUrl "https://myapplications.microsoft.com" -SendInvitationMessage:$true
    }
    # Update Account Created column on sharepoint list so this account is skipped in the future
    set-pnplistitem -Identity $termaccount.id -Values @{'IsCreated'=$true} -List $listname
}
