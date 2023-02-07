<#                   
.SYNOPSIS
    Change Employee Manager from SharePoint List
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
$sharepoint = "FORM SITE SHAREPOINT URL"
$listName = "EmploymentChange"
# ID of application in Azure
$ClientID = "ID OF APP INSTALLED IN AZURE"
# Thumbprint of certificate for authentication installed from setup process
$thumbprint = "CERTIFICATE THUMBPRINT INSTALLED ON MACHINE"
$cert = Get-ChildItem Cert:\LocalMachine\My\$Thumbprint
# 365 and on prem info
$Tenant = 'ORG.ONMICROSOFT.COM'
$tenantid = "TENANT ID FROM AZURE PROPERTIES"
#########################
# Connect to SharePoint #
#########################
# Gather new hire form information
# logs into sharepoint with certificate and App registered in Azure
Connect-PnPOnline -url $sharepoint -ClientId $ClientID -Tenant $Tenant -Thumbprint $thumbprint
Connect-MgGraph -ClientId $ClientID -TenantId $Tenantid -Certificate $cert
# Getting context
$useraccounts = Get-PnPListItem -List $listname | where-object { $_.FieldValues.ManagerChangeDone -ne $true -and $_.FieldValues.ManagerChange -eq $true }
##############
# Run Script #
##############
$date = get-date
ForEach ($UserAccount in $UserAccounts) {
    # Is this the day of the change
    if ($UserAccount.FieldValues.EffectiveDate -lt $date) {
        $Email = $null; $ManagerEmail = $null; $Manager = $null; $aaduser = $null
        $Email = $useraccount.FieldValues.Person.Email
        # Manager Change
        if ($UserAccount.FieldValues.ManagerChange -eq $true) {
            if ($null -ne $UserAccount.FieldValues.Person) {
                $ManagerEmail = $useraccount.FieldValues.NewManager.Email
                $Manager = Get-ADUser -Filter { UserPrincipalName -eq $ManagerEmail }
                $aaduser = Get-mguser -UserID $UPN -property OnPremisesSyncEnabled
                $aznewmanager = get-mguser -userid $newmanageremail
                if ($aaduser.OnPremisesSyncEnabled -eq $true) {
                    Get-ADUser -Filter { UserPrincipalName -eq $Email } | Set-ADUser -Manager $Manager
                } else {
                    $managerid = $aznewmanager.id
                    $params = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$managerid"}
                    Set-MgUserManagerByRef -userid $email -bodyparameter $params
                }
                set-pnplistitem -Identity $UserAccount.id -Values @{'ManagerChangeDone'=$true} -List $listname
            }
        }
    }
}
