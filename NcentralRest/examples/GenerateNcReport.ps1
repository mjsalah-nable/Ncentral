<#
.SYNOPSIS
    Generate a report of the NCs in the current directory.

.DESCRIPTION
    This script will generate a report of the NCs in the current directory. The report will be saved in the current directory as "NCReport.txt".

.AUTHOR
    Written by: Mohamed Salah
    Date: 2024-02-28
#>

Import-Module ./NcentralRest/NcentralRest.psm1

# Define the URL and token values for the API call
$ApiHost = "https://api.example.com/ncs"
$jwt = ""

Connect-Ncentral -ApiHost $ApiHost -Key $jwt

# Get the list of Organization Units, and save the list to a Map collection using orgUnitId as the key
$page = 1
$pageSize = 500
$continue = $true

do {
    $orgUnits = Get-NcentralOrgUnits -Page $page -PageSize $pageSize
    foreach ($orgUnit in $orgUnits) {
        $orgUnitMap[$orgUnit.orgUnitId] = $orgUnit
    }
    $page++
    $continue = $orgUnits.Count -eq $pageSize
} while ($continue)

# Next, let's get the list of devices from the NC. devices will be aggregated based on the orgUnitId and saved to a Map collection
# The collection will then be used to generate CSV report
$page = 1
$pageSize = 500
$continue = $true
$aggregatedReport = @{}

do {
    $devices = Get-NcentralDevices -Page $page -PageSize $pageSize
    foreach ($device in $devices) {
        $orgUnitId = $device.customerId
        $aggregatedStruct = $aggregatedReport[$orgUnitId]
        if ($null -eq $aggregatedStruct) {
            $BusinessUnit = $orgUnitMap[$orgUnitId]

            # Resolve so_id, so_name, customer_id, customer_name, site_id, site_name depending on $BusinessUnit's orgUnitType (SO/CUSTOMER/SITE)
            if ($BusinessUnit.orgUnitType -eq "CUSTOMER") {
                $SoUnit = $orgUnitMap[$BusinessUnit.parentId]
                $Customer = $BusinessUnit
                $Site = $null
            }
            elseif ($BusinessUnit.orgUnitType -eq "SITE") {
                $Site = $BusinessUnit
                $Customer = $orgUnitMap[$BusinessUnit.parentId]
                $SoUnit = $orgUnitMap[$Segment.parentId]
            }


            $aggregatedStruct = [PSCustomObject]@{
                'N-Able ID'          = $Customer.orgUnitId
                'Segment'            = $SoUnit.orgUnitName 
                'Business Unit Name' = $BusinessUnit.orgUnitName
                #'BU Windows Assets Discovered' = $BusinessUnitDeviceCount
                'BU Total Assets'    = 0
                #'BU Probe Count' = Get-ProbeCount($BusinessUnit.CustomerID)
                #'BU Discovery Started Date' =  Get-DiscoveryStartedDate($BusinessUnitProbe)
                'Site ID'            = if ($Site) { $Site.orgUnitId } else { "" }
                'Site Name'          = if ($Site) { $Site.orgUnitName } else { "" }
                #'Site Windows Assets Discovered' = $SiteDeviceCount
                'Site Total Assets'  = 0
                #'Site Probe Count' = Get-ProbeCount($Site.CustomerID)
                #'Site Discovery Started Date' =  Get-DiscoveryStartedDate($SiteProbe)
            }

            $aggregatedReport[$orgUnitId] = $aggregatedStruct
        }

        $aggregatedStruct.'BU Total Assets'++
        if ($aggregatedReport.'Site ID') {
            $aggregatedStruct.'Site Total Assets'++
        }
    }
    $page++
    $continue = $devices.Count -eq $pageSize
} while ($continue)


# When done, disconnect from the API server
Disconnect-Ncentral