<#
    Jake Cross
    9-24-22
    Copy HubSite top nav bar to other hub sites
#>
# Should be home site for tenant
$HomeSite = "https://domain.sharepoint.com"
Import-Module -Name PnP.PowerShell
Connect-PnPOnline -UseWebLogin -Url $HomeSite
# Get a list of hubs that aren't the home site
$hubs = Get-PnPHubSite | Where-Object { $_.SiteUrl -ne $HomeSite }
# Get top level links in menu tree on the home site
$navnodes = Get-PnPNavigationNode -Location TopNavigationBar
Disconnect-PnPOnline
# For every hub that isn't the home site add sharepoint top menu links
foreach ($hub in $hubs) {
    Connect-PnPOnline -UseWebLogin -Url $hub.SiteUrl
    # Remove whatever is currently on the destination hub site hub nav bar
    Get-PnPNavigationNode -Location TopNavigationBar | Remove-PnPNavigationNode -Force
    # For every top level link in menu tree
    foreach ($navnode in $navnodes) {
        Connect-PnPOnline -UseWebLogin -Url $HomeSite
        # Get object of current top level link to be able to grab children from it later
        $navnodeobject = Get-PnPNavigationNode -Id $navnode.id
        Disconnect-PnPOnline
        Connect-PnPOnline -UseWebLogin -Url $hub.SiteUrl
        # Create link on the destination hub site and store it in an object for later
        $topmenuobject = Add-PnPNavigationNode -Location TopNavigationBar -Title $navnode.Title -Url $navnode.url -External
        Disconnect-PnPOnline
        # For every second level link under the current top level link
        foreach ($child in $navnodeobject.children) {
            Connect-PnPOnline -UseWebLogin -Url $HomeSite
            # Get object of current second level link to be able to grab children from it later
            $childobject = Get-PnPNavigationNode -Id $child.id
            Disconnect-PnPOnline
            Connect-PnPOnline -UseWebLogin -Url $hub.SiteUrl
            # Create link on destination hub site and store it in an object for later
            $secondmenuobject = Add-PnPNavigationNode -Location TopNavigationBar -Title $childobject.Title -Url $childobject.url -External -Parent $topmenuobject.id
            # For every third level link under the current second level link
            foreach ($grandchild in $childobject.children) {
                # Create link on the destination hub site
                Add-PnPNavigationNode -Location TopNavigationBar -Title $grandchild.Title -Url $grandchild.url -External -Parent $secondmenuobject.id
            }
        }
    }
}