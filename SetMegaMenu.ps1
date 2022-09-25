<#
	Jake Cross
    9-24-22
    Set horizontal nav and mega menu everywhere
#>
Import-Module PnP.PowerShell
Connect-PnPOnline -UseWebLogin -Url "https://domain.sharepoint.com"
# Get all sites in tenant
$Sites = Get-PnPTenantSite
foreach ($site in $sites) {
    # Set user as secondary admin on all sites without adding them to a team
    Set-PnPTenantSite -Url $site.url -owners "user@domain.com"
}
Disconnect-PnPOnline
# Activate Communication Site style navigation on all sites
foreach ($Site in $Sites) {
    Connect-PnPOnline -UseWebLogin -Url $Site.url
    $Web = Get-PnPWeb
    $web.MegaMenuEnabled = $true
    $Web.HorizontalQuickLaunch = $true
    $Web.Update()
    Invoke-PnPQuery
    Disconnect-PnPOnline
}
