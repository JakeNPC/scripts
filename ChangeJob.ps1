<#                   
.SYNOPSIS
    Job Change from SharePoint List
.NOTES
    Authors:    Jake Cross
    Email:      jake@cross.house
    Updated:    2-6-23
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
$formsite = "FORM SITE URL"
$listName = "EmploymentChange"
# ID of application in Azure
$ClientID = "INSERT CLIENT ID OF APP REGISTERED IN AZURE"
# Thumbprint of certificate for authentication installed from setup process
$Thumbprint = "INSERT CERTIFICATE THUMBPRINT FOR AUTH"
$cert = Get-ChildItem Cert:\LocalMachine\My\$Thumbprint
# 365 tenant
$Tenant = '_____ORGNAME____.onmicrosoft.com'
$tenantid = "INSERT TENANT ID FROM AZURE TENANT PROPERTIES"
$defaultPW = "DEFAULT PW"
$oupath = "DN OF DEFAULT OU FOR USERS"
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
# For adding users to azure apps
#########################
# Connect to SharePoint #
#########################
# Gather new hire form information
# logs into sharepoint with certificate and App registered in Azure
Connect-PnPOnline -url $sharepoint -ClientId $ClientID -Tenant $Tenant -Thumbprint $thumbprint
# Connect to azure ad
Connect-MgGraph -ClientId $ClientID -TenantId $Tenantid -Certificate $cert
################################
# Collect info from SharePoint #
################################
# Grab job changes that have been approved by IT but haven't been completed
$useraccounts = get-pnplistitem -List $listname | where-object { $_.FieldValues.ApprovedByIT -eq $true -and $_.FieldValues.JobChange -eq $true -and $_.FieldValues.JobChangeDone -ne $true }
##############
# Run Script #
##############
# Get todays date
$date = get-date
# Loop through every account with a pending job change
ForEach ($UserAccount in $UserAccounts) {
    # If we have passed the day of change
    if ($UserAccount.FieldValues.EffectiveDate -lt $date) {
        # clears values from previous loops
        $LicensesToAssign = $null; $sam = $null; $plans = $null
        # Grabs the job title of the current Change
        $JobTitleObject = Get-PnPListItem -List 'Job Titles' -id $useraccount.FieldValues.Job_x0020_Title.LookupID
        $JobTitlegroupids = $JobTitleObject.FieldValues.Assignments.LookupID
        # Grabs the department of the current new hire from the job title
        $DepartmentObject = Get-PnPListItem -List 'Departments' -id $JobTitleObject.FieldValues.Department.LookupID
        $Departmentgroupids = $DepartmentObject.FieldValues.Assignments.LookupID
        # Get string of new job title of the subject of change
        $JobTitleValue = $UserAccount.FieldValues.New_x0020_Title.LookupValue
        # Get Azure Ad User Object of the subject of change
        $AzUser = get-mguser -userid $useraccount.FieldValues.Person.Email -property OnPremisesSyncEnabled, JobTitle, UserPrincipalName, id, GivenName, Surname, Department, OfficeLocation, City, State, StreetAdress, PostalCode, CompanyName
        # Get the UPN of the subject of change
        $UPN = $AzUser.UserPrincipalName
        # Add to Azure groups based on department and job title
        $JobTitleAzureGroups = get-assignmentsbytype -groupids $JobTitlegroupids -AssignmentType "Azure Group"
        $DepartmentAzureGroups = get-assignmentsbytype -groupids $Departmentgroupids -AssignmentType "Azure Group"
        Add-MemberToGroups -GroupType "Azure" -Groups $JobTitleAzureGroups -AADUserId $azuser.Id
        Add-MemberToGroups -GroupType "Azure" -Groups $DepartmentAzureGroups -AADUserId $azuser.Id
        ###########
        # License #
        ###########
        # capture current license info
        $oldlicenses = get-mguserlicenseDetail -UserId $UPN
        # Set office license based off of job title
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
        Set-MgUserLicense -UserId $UPN -AddLicenses $LicensesToAssign -RemoveLicenses @($oldlicenses)
        ######################
        # Account and Groups #
        ######################
        # If the account is syncronized with the local directory update the job title in AD and add to groups associated with new position and department
        If ($AzUser.OnPremisesSyncEnabled -ne $True) {
            If ($JobTitleID.FieldValues.OnPremAccount -eq $true) {
                # setting account fields based on current attributes
                $dn = $AzUser.DisplayName ; $FirstName = $AzUser.GivenName; $LastName = $AzUser.SurName; $Department = $AzUser.Department; $Company = $AzUser.CompanyName
                $JobTitle = $AzUser.JobTitle; $Office = $AzUser.OfficeLocation; $city = $AzUser.City; $State = $AzUser.State; $Zip = $AzUser.PostalCode ; $Address = $AzUser.StreetAddress
                # Grabbing current manager info - this will break if the manager doesn't have an on prem account - may need to fix this later
                $365Managerid = get-mgusermanager -userid $azuser.id
                $Manager = Get-ADUser -Filter { UserPrincipalName -eq $365Manager.AdditionalProperties.UserPrincipalName }
                # set sam account name
                $Sam = $azuser.UserPrincipalName.Split('@')[0]
                # make an on prem account
                New-ADuser -SamAccountName $sam -GivenName $FirstName -Surname $LastName -DisplayName $dn -Name $dn -accountPassword (ConvertTo-SecureString $defaultPW -AsPlainText -Force) -UserPrincipalName $UPN `
                    -EmailAddress $UPN -Path $OUpath -Department $Department -Company $Company -Manager $Manager.Samaccountname -Title $JobTitle -Description $JobTitle -Office $Office -City $City -State $State `
                    -PostalCode $Zip -StreetAddress $Address -Country "US" -Enabled $true -ChangePasswordAtLogon $true
                # Adding proxy addresses to account to stitch the on prem account to the azure account
                $ProxyAddresses = ('SMTP:' + $UPN), ('smtp:' + $sam + '@' + $tenant)
                Set-ADUser -identity $sam -Add @{'proxyAddresses' = $ProxyAddresses }
                # Sync AD with 365 and wait
                start-ADSyncSyncCycle -policytype delta
                start-sleep 120
            }
        }
        # refresh Azure Ad User Object of the subject of change
        $AzUser = get-mguser -userid $useraccount.FieldValues.Person.Email -property OnPremisesSyncEnabled
        # Now that on prem accounts have been created if there weren't already try adding account to on prem groups
        if ($AzUser.OnPremisesSyncEnabled -eq $True) {
            # Get AD account of subject of change by their UPN
            $jobtitleuser = Get-ADUser -Filter { UserPrincipalName -eq $UPN }
            # Set the job title on the AD account of the subject of change to their new job title
            Set-ADUser $jobtitleuser -Title $JobTitleValue -Department $DepartmentID.FieldValues.Title -Description $JobTitleValue
            # setting sam variable to be used in group functions
            $sam = $jobtitleuser.Samaccountname
            # Adds user to groups based on department
            $DepartmentADGroups = get-assignmentsbytype -groupids $Departmentgroupids -AssignmentType "AD Group"
            Add-MemberToGroups -groups $DepartmentADGroups -GroupType "AD"
            # Adds user to groups based on job title
            $JobTitleADGroups = get-assignmentsbytype -groupids $JobTitlegroupids -AssignmentType "AD Group"
            Add-MemberToGroups -groups $JobTitleADGroups -GroupType "AD"
        }
        set-pnplistitem -Identity $UserAccount.id -Values @{'JobChangeDone'=$true} -List $listname
    }
}
