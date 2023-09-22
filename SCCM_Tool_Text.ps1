function connectSCCM {
	$server = "WAPSCM01094P03.corp.costco.com"
	$site = "CM1"

	Set-Location 'C:\Program Files (x86)\ConfigMgr\bin\'
	Import-Module .\ConfigurationManager.psd1

	If (!(Test-Path CM1:)) {
		New-PSDrive -Name $site -PSProvider "CMSite" -Root $server -Description "SCCM site"
        Write-Host "Created CM1 PSDrive..."
	}
	Set-Location CM1:
}

function getDevice {
	param([string]$hostname)
	$server = "WAPSCM01094P03.corp.costco.com"
	$site = "CM1"

	$computer = Get-CMDevice -Name $hostname -CollectionId CM101624 -Resource
	$collections = (Get-WmiObject -ComputerName $server -Namespace ("root/SMS/Site_"+$site) -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$hostname' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID").Name 

	$collectionArray = $collections.Split("`n")
	$collectionsArray = $collectionsArray | Sort-Object 

	$computerObj = [pscustomobject]@{
		Name = ($computer.Name).ToUpper()
		#OU = $computer.DistinguishedName
		IPv4 = $computer.IPAddresses[0]
		ResourceID = $computer.ResourceId
		Collections = $collectionArray
	}
	
	return $computerObj
}

function hostnamePrompt {
	param([string]$msg)
	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

	$title = 'SCCM Configuration Tool'
	$output = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
	return $output
}

function addToCollection {
	param([string]$resourceId,[string]$collection)

	try {
		Add-CMDeviceCollectionDirectMembershipRule -CollectionName $collection -resourceId $resourceId -ErrorAction Stop
		Write-Host "Successfully added device to collection ""$collection""." -BackgroundColor Black -ForegroundColor Green
		return $true
	} catch {
		Write-Host "Unable to add device to collection ""$collection""." -BackgroundColor Black -ForegroundColor Red
		return $false
	}
}

function popupBox {
	param([string]$msg,[string]$type)
	[System.Windows.Forms.MessageBox]::Show($msg,$type)
}

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

connectSCCM

Write-Host "========================================"
Write-Host "========================================"
Write-Host "========= SCCM Collection Tool ========="
Write-Host "========================================"
Write-Host "========================================"

$hostname1 = hostnamePrompt -msg "Enter name of computer with MISSING collections: "
$hostname2 = hostnamePrompt -msg "Enter name of computer with WORKING collections: "

$computerArray = @()
$collectionsArray = @()

$loopCounter = 0
while ($loopCounter -lt 5) {
	if ($loopCounter -gt 4) {
		popupBox -msg "Unable to retrieve information for device $hostname1. Script will now exit." -type 'WARNING'
		exit
	}

	$computer1 = getDevice -hostname $hostname1
	if ([string]::IsNullOrWhiteSpace($computer1)) {
		popupBox -msg "Unable to retrieve information for device $hostname1. Please check name and try again." -type 'WARNING'
        $location = Get-Location
        Write-Host "Current location is $location."
		$loopCounter++
		$hostname1 = hostnamePrompt -msg "Enter name of computer with MISSING collections: "
	} else {
		$computerArray += $computer1
		break
	}
}

while ($loopCounter -lt 5) {
	if ($loopCounter -gt 4) {
		popupBox -msg "Unable to retrieve information for device $hostname1. Script will now exit." -type 'WARNING'
		exit
	}

	$computer2 = getDevice -hostname $hostname2
	if ([string]::IsNullOrWhiteSpace($computer1)) {
		popupBox -msg "Unable to retrieve information for device $hostname1. Please check name and try again." -type 'WARNING'
        $location = Get-Location
        Write-Host "Current location is $location."
		$loopCounter++
		$hostname2 = hostnamePrompt -msg "Enter name of computer with WORKING collections: "
	} else {
		$computerArray += $computer2
		break
	}
}

:outer
foreach ($collection2 in $computer2.Collections) {
	:inner
	foreach ($collection1 in $computer1.Collections) {
		if ($collection2 -eq $collection1) {
			$collectionObj = [pscustomobject]@{
				$hostname1 = $collection1
				$hostname2 = $collection2
			}
			$collectionsArray += $collectionObj
			continue outer
		}
	}
	$collectionObj = [pscustomobject]@{
		$hostname1 = "NOT FOUND - " + $collection2
		$hostname2 = $collection2
	}
	$collectionsArray += $collectionObj
}

foreach ($computer in $computerArray) {
	Write-Host $computer.Name + "'s information:"
	Write-Host "IPv4 address: " + $computer.IPv4
	Write-Host "Resource ID: " + $computer.ResourceID
}

$collectionsArray | Format-Table -AutoSize

#$checkArray = @()

#:outer
#foreach ($item2 in $collections2) {
#	:inner
#	foreach ($item1 in $collections1) {
#		if ($item2 -eq $item1) {
#			continue outer
#		}
#	}
#	$addResult = addToCollection -resourceId $computer1.ResourceID -collection $item2
#	$checkArray += [pscustomobject]@{
#		Collection = $item2
#		Result = $addResult
#	}
#}

#$exportPath = "C:\Temp\SCCM_Collection_Results.csv"
#$checkArray | Sort-Object Result | Export-Csv -Path $exportPath -Force
#Write-Host "Exported results to $exportPath."
#$checkArray | Sort-Object Result | Format-Table -AutoSize

Read-Host "Press Enter to exit"