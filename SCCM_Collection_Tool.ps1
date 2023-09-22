#SCCM CONFIG TOOL - Please contact khoaphan@costco.com if there are questions or concerns.

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# ADMIN ELEVATION, SECURITY PERMISSION CHECK AND VARIOUS FUNCTIONS

# ELEVATE TO ADMIN IF SCRIPT WAS RAN AS NORMAL USER
param(
    [Parameter(Mandatory=$false)]
    [switch]$shouldAssumeToBeElevated,

    [Parameter(Mandatory=$false)]
    [String]$workingDirOverride
)

# If parameter is not set, we are propably in non-admin execution. We set it to the current working directory so that
#  the working directory of the elevated execution of this script is the current working directory
if(-not($PSBoundParameters.ContainsKey('workingDirOverride')))
{
    $workingDirOverride = (Get-Location).Path
}

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# If we are in a non-admin execution. Execute this script as admin
if ((Test-Admin) -eq $false)  {
    if ($shouldAssumeToBeElevated) {
        Write-Output "Elevating did not work :("

    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -shouldAssumeToBeElevated -workingDirOverride "{1}"' -f ($myinvocation.MyCommand.Definition, "$workingDirOverride"))
    }
    exit
}

Set-Location "$workingDirOverride"

# END ADMIN ELEVATION
# Add actual commands to be executed in elevated mode here:

# CHANGE SECURITY PERMISSIONS FOR SCCM CONNECTION
function changeSecPolicy {
    param([string]$policyName,[string]$policyFullName)
	$accountToAdd = "costco\$Env:UserName" #Retrieves username of current user
    $sidstr = $null
    try {
	    $ntprincipal = new-object System.Security.Principal.NTAccount "$accountToAdd"
	    $sid = $ntprincipal.Translate([System.Security.Principal.SecurityIdentifier])
	    $sidstr = $sid.Value.ToString()
    } catch {
	    $sidstr = $null
    }
    Write-Host "Account: $($accountToAdd)" -ForegroundColor DarkCyan
    if( [string]::IsNullOrEmpty($sidstr) ) {
	    Write-Host "Account not found!" -ForegroundColor Red
	    exit -1
    }
    Write-Host "Account SID: $($sidstr)" -ForegroundColor DarkCyan
    $tmp = ""
    $tmp = [System.IO.Path]::GetTempFileName()
    Write-Host "Exporting current Local Security Policy..." -ForegroundColor DarkCyan
    secedit.exe /export /cfg "$($tmp)" 
    $c = ""
    $c = Get-Content -Path $tmp
    $currentSetting = ""
    foreach($s in $c) {
	    if( $s -like $policyName) {
		    $x = $s.split("=",[System.StringSplitOptions]::RemoveEmptyEntries)
		    $currentSetting = $x[1].Trim()
	    }
    }
    if( $currentSetting -notlike "*$($sidstr)*" ) {
	    Write-Host "Modifying Setting ""$policyFullName""" -ForegroundColor DarkCyan
	
	    if( [string]::IsNullOrEmpty($currentSetting) ) {
		    $currentSetting = "*$($sidstr)"
	    } else {
		    $currentSetting = "*$($sidstr),$($currentSetting)"
	    }
	
	    Write-Host "$currentSetting"
	
	$outfile = @"
[Unicode]
Unicode=yes
[Version]
signature="`$COSTCOWHOLESALE`$"
Revision=1
[Privilege Rights]
$policyName = $($currentSetting)
"@
	
	    $tmp2 = ""
	    $tmp2 = [System.IO.Path]::GetTempFileName()	
	
	    Write-Host "Importing new settings to Local Security Policy..." -ForegroundColor DarkCyan
	    $outfile | Set-Content -Path $tmp2 -Encoding Unicode -Force
	    Push-Location (Split-Path $tmp2)
	
	    try {
		    secedit.exe /configure /db "secedit.sdb" /cfg "$($tmp2)" /areas USER_RIGHTS
	    } finally {	
		    Pop-Location
	    }
    } else {
	    Write-Host "NO ACTIONS REQUIRED! Account is already in ""$policyFullName""." -ForegroundColor DarkCyan
    }
    Write-Host "Done.`r`n" -ForegroundColor DarkCyan
}

# CONNECT TO SCCM VIA PSDRIVE
function connectSCCM {
    Write-Host "Connecting to SCCM... Please feel free to ignore or minimize this window."
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

# GET DEVICE INFO FROM SCCM
function getDevice {
	param([string]$hostname)
	$server = "WAPSCM01094P03.corp.costco.com"
	$site = "CM1"

    try {
	    $computer = Get-CMDevice -Name $hostname -Resource -ErrorAction Stop
    } catch {
        $computerObj = [pscustomobject]@{
		    Name = "Unable to retrieve device."
		    #OU = $computer.DistinguishedName
		    IPv4 = "Unable to retrieve device."
		    ResourceID = "Unable to retrieve device."
		    Collections = "Unable to retrieve device."
        }
        return $computerObj
    }

	$collections = (Get-WmiObject -ComputerName $server -Namespace ("root/SMS/Site_"+$site) -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$hostname' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID").Name 

    if ([string]::IsNullOrWhiteSpace($computer.Name) -or [string]::IsNullOrWhiteSpace($collections)) {
         $computerObj = [pscustomobject]@{
		    Name = "Unable to retrieve device."
		    #OU = $computer.DistinguishedName
		    IPv4 = "Unable to retrieve device."
		    ResourceID = "Unable to retrieve device."
		    Collections = "Unable to retrieve device."
        }
        return $computerObj
	}

	$collectionArray = $collections.Split("`n")
	$collectionArray = $collectionArray | Sort-Object 

	$computerObj = [pscustomobject]@{
		Name = ($computer.Name).ToUpper()
		#OU = $computer.DistinguishedName
		IPv4 = $computer.IPAddresses[0]
		ResourceID = $computer.ResourceId
		Collections = $collectionArray
	}
	
	return $computerObj
}

# ADD DEVICE TO A COLLECTION
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

# REMOVE DEVICE FROM A COLLECTION
function removeFromCollection {
	param([string]$resourceId,[string]$collection)

	try {
		Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $collection -resourceId $resourceId -Confirm False -Force -ErrorAction Stop
		Write-Host "Successfully removed device from collection $collection."
		return $true
	} catch {
		Write-Host "Unable to remove device to collection $collection."
		return $false
	}
}

#Triggers SCCM actions to refresh Software Center
function Invoke-ConfigMgrClientAction
{
[cmdletbinding()]
Param
(
    [Parameter(Mandatory=$False,ValueFromPipeline=$True)]
    [string[]]$ComputerName = "localhost",

    [Parameter(Mandatory=$False,ValueFromPipeline=$False)]
    [ValidateSet('MachinePolicies','SoftwareUpdateEvaluationCycle','ApplicationPolicy')]
    [string[]]$Actions = @('MachinePolicies','SoftwareUpdateEvaluationCycle','ApplicationPolicy')
)

    Begin
    {
        $ClientActions = @{
            'MachinePolicies'='{00000000-0000-0000-0000-000000000022}';
            'ApplicationPolicy'='{00000000-0000-0000-0000-000000000121}';
            'SoftwareUpdateEvaluationCycle'='{00000000-0000-0000-0000-000000000108}';
        }

        $jobs = @()
    }

    Process
    {
        $Credential = Get-Credential -Message "Please enter your LAN credentials to authenticate with SCCM."
        $jobs += Invoke-Command -ComputerName $ComputerName -Credential:$Credential -Authentication Default -ScriptBlock {
            $VerbosePreference = $using:VerbosePreference
            $ClientActions = $using:ClientActions
            $Actions = $using:Actions

            try
            {
                foreach ($action in $Actions)
                {
                    Write-Verbose "Triggering SCCM client action: '$action' with schedule: '$($ClientActions.$action)'"
                    Invoke-WMIMethod -Namespace "Root\CCM" -Class SMS_CLIENT -Name TriggerSchedule -ArgumentList $ClientActions.$action -ErrorAction Stop | Out-String | Write-Debug
                }

                return "Successfully refreshed Software Center on computer $env:COMPUTERNAME."
            }
            catch
            {
                return "[$env:COMPUTERNAME] ERROR: $_"
            }
        } -AsJob
    }

    End
    {
        $jobs | Receive-Job -Wait
    }
}

# CLEAR ALL FIELDS ON GUI
function clearAll {
	$boxCollection1.Clear()
	$boxCollection2.Clear()
	$labelIP1.Text = "Detected IP: "
	$labelIP2.Text = "Detected IP: "
	$labelResults.Text = ""
}

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# MAIN SCRIPT

# SECURITY PERMISSIONS FOR SCCM CONNECTION
Write-Host "Checking security permissions..."
changeSecPolicy -policyName "SeBatchLogonRight*" -policyFullName "Log On as Batch Job"
changeSecPolicy -policyName "SeServiceLogonRight*" -policyFullName "Log On as Service"

# CONNECTS TO SCCM MODULE
connectSCCM

# GUI PACKAGES
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# FLAT/MODERN UI STYLE
[System.Windows.Forms.Application]::EnableVisualStyles()

# BASE FORM
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM Collection Tool - Main'
$form.Size = New-Object System.Drawing.Size(800,670)
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Add_Load({
	$form.Activate()
})

# LABELS, BUTTONS AND TEXT BOXES
$labelComp1 = New-Object System.Windows.Forms.Label
$labelComp1.Location = New-Object System.Drawing.Size(80,5)
$labelComp1.Size = New-Object System.Drawing.Size(300,30)
$labelComp1.Text = 'Computer with MISSING software:'
$labelComp1.Font = ‘Segoe UI,12’
$labelComp1.ForeColor = [System.Drawing.Color]::Black
$labelComp1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelComp1)

$boxComp1 = New-Object System.Windows.Forms.TextBox
$boxComp1.Location = New-Object System.Drawing.Point(10,35)
$boxComp1.Size = New-Object System.Drawing.Size(380,30)
$boxComp1.Text = " "
$boxComp1.AutoSize = $false
$boxComp1.Font = ‘Segoe UI,11’
$boxComp1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxComp1)

$labelIP1 = New-Object System.Windows.Forms.Label
$labelIP1.Location = New-Object System.Drawing.Size(10,80)
$labelIP1.Size = New-Object System.Drawing.Size(100,25)
$labelIP1.Text = 'Detected IP: '
$labelIP1.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$labelIP1.ForeColor = [System.Drawing.Color]::Black
$labelIP1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelIP1)

$IPaddr1 = New-Object System.Windows.Forms.Label
$IPaddr1.Location = New-Object System.Drawing.Size(110,80)
$IPaddr1.Size = New-Object System.Drawing.Size(200,25)
$IPaddr1.Text = ''
$IPaddr1.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$IPaddr1.ForeColor = [System.Drawing.Color]::Black
$IPaddr1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($IPaddr1)

$boxCollection1 = New-Object System.Windows.Forms.RichTextBox
$boxCollection1.Location = New-Object System.Drawing.Point(10,115)
$boxCollection1.Size = New-Object System.Drawing.Size(380,400)
$boxCollection1.AutoSize = $false
$boxCollection1.Multiline = $true
$boxCollection1.WordWrap = $true
$boxCollection1.Font = ‘Segoe UI,11’
$boxCollection1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxCollection1)

$labelComp2 = New-Object System.Windows.Forms.Label
$labelComp2.Location = New-Object System.Drawing.Size(460,5)
$labelComp2.Size = New-Object System.Drawing.Size(300, 30)
$labelComp2.Text = 'Computer with CORRECT software:'
$labelComp2.Font = ‘Segoe UI,12’
$labelComp2.ForeColor = [System.Drawing.Color]::Black
$labelComp2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelComp2)

$boxComp2 = New-Object System.Windows.Forms.TextBox
$boxComp2.Location = New-Object System.Drawing.Point(400,35)
$boxComp2.Size = New-Object System.Drawing.Size(370,30)
$boxComp2.Text = " "
$boxComp2.AutoSize = $false
$boxComp2.Font = ‘Segoe UI,11’
$boxComp2.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxComp2)

$labelIP2 = New-Object System.Windows.Forms.Label
$labelIP2.Location = New-Object System.Drawing.Size(400,80)
$labelIP2.Size = New-Object System.Drawing.Size(100,25)
$labelIP2.Text = 'Detected IP: '
$labelIP2.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$labelIP2.ForeColor = [System.Drawing.Color]::Black
$labelIP2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelIP2)

$IPaddr2 = New-Object System.Windows.Forms.Label
$IPaddr2.Location = New-Object System.Drawing.Size(500,80)
$IPaddr2.Size = New-Object System.Drawing.Size(200,25)
$IPaddr2.Text = ''
$IPaddr2.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$IPaddr2.ForeColor = [System.Drawing.Color]::Black
$IPaddr2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($IPaddr2)

$boxCollection2 = New-Object System.Windows.Forms.RichTextBox
$boxCollection2.Location = New-Object System.Drawing.Point(400,115)
$boxCollection2.Size = New-Object System.Drawing.Size(370,400)
$boxCollection2.AutoSize = $false
$boxCollection2.Multiline = $true
$boxCollection2.WordWrap = $true
$boxCollection2.Font = ‘Segoe UI,11’
$boxCollection2.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxCollection2)

# FIX BUTTON
$fix = {
	$labelResults.Text = "Launching new window... When done, please WAIT 30 SECONDS and click COMPARE again."

	[array]$tmpArray = $boxCollection1.Text.Split("`n")
	[array]$badCollections = @()
	foreach ($collection in $tmpArray) {
		if ($collection.StartsWith("NOT FOUND")) {
			$collection = $collection.Substring(12)
			$badCollections += $collection
		}
	}

	$PSScriptRoot
	& "$PSScriptRoot\fixCollections.ps1" -computerName $boxComp1.Text -collectionArray $badCollections
}

$buttonFix = New-Object System.Windows.Forms.Button
$buttonFix.Location = New-Object System.Drawing.Point(130,540)
$buttonFix.Size = New-Object System.Drawing.Size(170,30)
$buttonFix.Text = 'BUTTON DISABLED'
$buttonFix.Font = ‘Segoe UI,12’
$buttonFix.ForeColor = [System.Drawing.Color]::White
$buttonFix.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonFix.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonFix.BackColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonFix.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkGreen
$buttonFix.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Green
$form.Controls.Add($buttonFix)
$buttonFix.Add_Click($fix) #Adds function to button
$buttonFix.Enabled = $false

# COMPARE BUTTON
$compare = {	
    clearAll
    $PC1Name = ($boxComp1.Text).Trim().ToUpper()
    $PC2Name = ($boxComp2.Text).Trim().ToUpper()

	if ($PC1Name.StartsWith("W10W") -and $PC2Name.StartsWith("W10W")) {
		#$location = Get-Location
		$labelResults.Text = "Requesting data from SCCM, this may take up to 1 minute depending on your network..."

		$computer1 = getDevice -hostname $PC1Name
		$computer2 = getDevice -hostname $PC2Name

		if ($computer1.IPv4 -eq "Unable to retrieve device.") {
			$IPaddr1.ForeColor = 'OrangeRed'
			$IPaddr1.Text = $computer1.IPv4
		} else {
			$IPaddr1.ForeColor = 'Green'
			$IPaddr1.Text = $computer1.IPv4
		}

		if ($computer2.IPv4 -eq "Unable to retrieve device.") {
			$IPaddr2.ForeColor = 'OrangeRed'
			$IPaddr2.Text = $computer2.IPv4
		} else {
			$IPaddr2.ForeColor = 'Green'
			$IPaddr2.Text = $computer2.IPv4
		}		
		
		$boxComp1.Text = " "+$computer1.Name
		$boxComp2.Text = " "+$computer2.Name

		$counter = 0
		:outer
		foreach ($collection2 in $computer2.Collections) {
			:inner
			foreach ($collection1 in $computer1.Collections) {
				if ($collection2 -eq $collection1) {
					$boxCollection1.AppendText($collection1+"`r`n")				
					$boxCollection2.AppendText($collection2+"`r`n")
					$counter++
					continue outer
				}
			}
			$boxCollection1.SelectionColor = 'orangered'
			$boxCollection1.AppendText("NOT FOUND - "+$collection2+"`r`n")	

			$boxCollection2.SelectionColor = 'royalblue'
			$boxCollection2.AppendText($collection2+"`r`n")
			$counter++
		}

		if ($counter -gt 10) {
			$boxCollection1.ScrollBars = "Vertical"
			$boxCollection2.ScrollBars = "Vertical"
		}

		$labelResults.Text = "Successfully retrieved data!"
		$buttonFix.Enabled = $true
		$buttonFix.Font = ‘Segoe UI,12’
		$buttonFix.Text = "FIX COLLECTIONS"

		$buttonRefreshSWC.Enabled = $true
		$buttonRefreshSWC.Font = 'Segoe UI,12'
		$buttonRefreshSWC.Text = "REFRESH ENDPOINT"

		if ($computer1.IPv4 -eq "Unable to retrieve device.") {
			$boxCollection1.Clear()
			$boxCollection1.Append_Text = "Unable to retrieve information for this device. Please check device name/network and try again."
			$labelResults.Text = "Unable to retrieve data for one or more computer(s), please check device name/network and try again."

			$buttonFix.Text = 'BUTTON DISABLED'
			$buttonFix.Font = ‘Segoe UI,12’
			$buttonFix.Enabled = $false
			$buttonRefreshSWC.Text = 'BUTTON DISABLED'
			$buttonRefreshSWC.Font = ‘Segoe UI,12’
			$buttonRefreshSWC.Enabled = $false
		}

		if ($computer2.IPv4 -eq "Unable to retrieve device.") {
			$boxCollection2.Clear()
			$boxCollection1.Append_Text = "Unable to retrieve information for this device. Please check device name/network and try again."
			$labelResults.Text = "Unable to retrieve data for one or more computer(s), please check device name/network and try again.."

			$buttonFix.Text = 'BUTTON DISABLED'
			$buttonFix.Font = ‘Segoe UI,12’
			$buttonFix.Enabled = $false
			$buttonRefreshSWC.Text = 'BUTTON DISABLED'
			$buttonRefreshSWC.Font = ‘Segoe UI,12’
			$buttonRefreshSWC.Enabled = $false
		}	
	} else {
		$labelResults.Text = "Device name is incorrect or not supported. Please try again."
	}    
}

$buttonCompare = New-Object System.Windows.Forms.Button
$buttonCompare.Location = New-Object System.Drawing.Point(320,540)
$buttonCompare.Size = New-Object System.Drawing.Size(150,30)
$buttonCompare.Text = 'COMPARE'
$buttonCompare.Font = ‘Segoe UI,12’
$buttonCompare.ForeColor = [System.Drawing.Color]::White
$buttonCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCompare.FlatAppearance.BorderColor = [System.Drawing.Color]::DeepSkyBlue
$buttonCompare.BackColor = [System.Drawing.Color]::DeepSkyBlue
$buttonCompare.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkBlue
$buttonCompare.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::DodgerBlue
$form.Controls.Add($buttonCompare)
$buttonCompare.Add_Click($compare) #Adds function to button

# REFRESH SWC BUTTON
$refreshSWC = {
	$PCName = ($boxComp1.Text).Trim()
	$labelResults.Text = Invoke-ConfigMgrClientAction -ComputerName $PCName
}

$buttonRefreshSWC = New-Object System.Windows.Forms.Button
$buttonRefreshSWC.Location = New-Object System.Drawing.Point(490,540)
$buttonRefreshSWC.Size = New-Object System.Drawing.Size(170,30)
$buttonRefreshSWC.Text = 'BUTTON DISABLED'
$buttonRefreshSWC.Font = ‘Segoe UI,12’
$buttonRefreshSWC.ForeColor = [System.Drawing.Color]::White
$buttonRefreshSWC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRefreshSWC.FlatAppearance.BorderColor = [System.Drawing.Color]::LightCoral
$buttonRefreshSWC.BackColor = [System.Drawing.Color]::LightCoral
$buttonRefreshSWC.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkRed
$buttonRefreshSWC.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Crimson
$form.Controls.Add($buttonRefreshSWC)
$buttonRefreshSWC.Add_Click($refreshSWC) #Adds function to button
$buttonRefreshSWC.Enabled = $false

$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Location = New-Object System.Drawing.Point(10,600)
$labelResults.Size = New-Object System.Drawing.Size(900,30)
$labelResults.Text = ""
$labelResults.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold+[System.Drawing.FontStyle]::Italic)
$labelResults.ForeColor = [System.Drawing.Color]::Black
$labelResults.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelResults)

# POPUP MESSAGE WHEN USER CLOSES APP
$form.Add_Closing({param($sender,$e)
    $result = [System.Windows.Forms.MessageBox]::Show(`
        "Are you sure you want to exit?", `
        "Close", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
		return
		exit
	}
	
	if ($result -eq [System.Windows.Forms.DialogResult]::No) {
		$e.Cancel = $true
	}
})

# WHAT HAPPENS WHEN USER PRESSES ENTER
$form.AcceptButton = $buttonCompare

# SHOW FORM
$Null = $form.ShowDialog()