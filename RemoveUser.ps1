<#                   
.SYNOPSIS
    Termination
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
$formsite = "FORM SITE URL"
$listName = "Terminations"
# ID of application in Azure
$ClientID = "INSERT CLIENT ID OF APP REGISTERED IN AZURE"
# Thumbprint of certificate for authentication installed from setup process
$Thumbprint = "INSERT CERTIFICATE THUMBPRINT FOR AUTH"
$cert = Get-ChildItem Cert:\LocalMachine\My\$Thumbprint
# 365 tenant
$Tenant = '_____ORGNAME____.onmicrosoft.com'
$tenantid = "INSERT TENANT ID FROM AZURE TENANT PROPERTIES"
# Time offset from UTC
$timeoffset = -8
#########################
# Connect to Office 365 #
#########################
# logs into sharepoint with certificate and app registered in Azure
$formsiteconnection = Connect-PnPOnline -url $formsite -ClientId $ClientID -Tenant $Tenant -Thumbprint $Thumbprint -returnconnection
# Connect to azure ad
Connect-MgGraph -ClientId $ClientID -TenantId $Tenantid -Certificate $cert
# Connect to exchange online
Connect-ExchangeOnline -AppId $ClientID -Organization $Tenant -CertificateThumbprint $thumbprint
################################
# Collect info from SharePoint #
################################
# Get all users in employment list
$termaccounts = Get-PnPListItem -connection $formsiteconnection -List $listname | where-object { $_.FieldValues.Terminated -ne $true -and $_.FieldValues.ApprovedByIT -eq $true }
#####################
# Terminate Account #
#####################
# Loops through terminated users converts to shared mailboxes
$date = get-date
foreach ($termaccount in $termaccounts) {
    $termdate = $termaccount.FieldValues.TermDateTime
    # Time Offset
    $termdate = $termdate.addhours($timeoffset)
    if ($termdate -lt $date) {
        $User = $null; $UPN = $null; $aduser = $null; $drs = $null; $newmanager = $null; $newmanageremail = $null; $aznewmanager = $null;
        # Grabs UPN from form
        $UPN = $termaccount.FieldValues.Person.Email
        # Get Azure AD User
        $User = Get-mguser -UserID $UPN -property OnPremisesSyncEnabled, JobTitle, id, DisplayName, DirectReports
        # grabbing license details to remove license in case account is getting restored for email retention
        $license = get-mguserlicenseDetail -UserId $UPN
        # Grant Access to OneDrive
        $manager = get-mgusermanager -userid $user.id
        $onedrive = Get-PnPUserProfileProperty -Account $user.UserPrincipalName
        Set-PnPTenantSite -Url $onedrive.PersonalUrl -Owners $manager.AdditionalProperties.UserPrincipalName
        # Assign Direct reports to new manager
        # get direct reports
        $drs = Get-MgUserDirectReport -UserId $upn
        # If direct reports are present
        if ($null -ne $drs) {
            # get email from user object field
            $newmanageremail = $termaccount.FieldValues.TransferDRsto.Email
            # get ad user from email
            $newmanager = get-aduser -Filter "userPrincipalName -like '$newmanageremail'"
            # get azure ad user from email
            $aznewmanager = get-mguser -userid $newmanageremail
            # process all direct reports
            foreach ($dr in $drs) {
                $druser = $null; $drupn = $null;
                # if on prem
                $drupn = $dr.additionalproperties.userPrincipalName
                $drazuser = get-mguser -userid $drupn -property OnPremisesSyncEnabled
                if ($drazuser.OnPremisesSyncEnabled -eq $true) {
                    # update new manager
                    $druser = get-aduser -Filter "userPrincipalName -like '$drupn'"
                    $druser | Set-AdUser -Manager $newmanager
                }
                # if not on prem
                else {
                    # update new manager
                    $managerid = $aznewmanager.id
                    $params = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$managerid"}
                    Set-MgUserManagerByRef -userid $drupn -bodyparameter $params
                }
            }
        }
        # if account is AD account synced with 365
        if ($User.OnPremisesSyncEnabled -eq $true) {
            # Delete AD user and sync with 365
            $aduser = get-aduser -Filter "userPrincipalName -like '$UPN'"
            Remove-Aduser $aduser -confirm:$false
            Start-ADSyncSyncCycle -PolicyType Delta
            start-sleep 120
            # retain email and clear account attributes
            if ($jobtitle.FieldValues.EmailRetention -eq $true) {
                $uid = $user.id
                # restoring cloud account after on prem account was deleted
                Restore-MgDirectoryDeletedItem -DirectoryObjectId $uid
                # clearing manager
                Remove-MgUserManagerByRef -UserId $UPN
                # clearing other attributes
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/$uid" -Body @{
                    Department = $null;
                    JobTitle = $null;
                    CompanyName = $null;
                    OfficeLocation = $null;
                    AccountEnabled = $false;
                }
                # convert to shared mailbox
                Set-Mailbox -Identity $UPN -Type Shared -hiddenfromaddresslistenabled:$true
                # Remove user licenses
                Set-MgUserLicense -UserId $UPN -AddLicenses @() -RemoveLicenses @($license.skuid)
            }
        }
        # if account is cloud only
        elseif ($User.DirSyncEnabled -ne $true) {
            # Delete 365 account
            Remove-mguser -userid $user.id
        }
        # Update sharepoint list to mark as terminated to skip item in future runs
        set-pnplistitem -Identity $termaccount.id -Values @{'Terminated'=$true} -List $listname -connection $formsiteconnection
    }
}
