param ([string]$computerName,[array]$collectionArray)
$ErrorActionPreference = 'SilentlyContinue'

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

# CONNECT TO SCCM VIA PSDRIVE
function connectSCCM {
    #Write-Host "Connecting to SCCM..."
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

	if ([string]::IsNullOrWhiteSpace($hostname)) {
		$labelResults.Text = "Device name is $hostname"
        $computerObj = [pscustomobject]@{
		    Name = "Unable to retrieve device."
		    #OU = $computer.DistinguishedName
		    IPv4 = "Unable to retrieve device."
		    ResourceID = "Unable to retrieve device."
        }
        return $computerObj
	}

	$computer = Get-CMDevice -Name $hostname -Resource

    $computerObj = [pscustomobject]@{
		Name = ($computer.Name).ToUpper()
		#OU = $computer.DistinguishedName
		IPv4 = $computer.IPAddresses[0]
		ResourceID = $computer.ResourceId
	}
	
	return $computerObj
}

# REMOVES WHITESPACE FROM COMPUTER NAME
$computerName = $computerName.Trim()

# GUI PACKAGES
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# FLAT/MODERN UI STYLE
[System.Windows.Forms.Application]::EnableVisualStyles()

# BASE FORM
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM Collection Tool - Fix Collections'
$form.Size = New-Object System.Drawing.Size(600,670)
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Add_Load({
	$form.Activate()
})

# LABELS, BUTTONS AND TEXT BOXES
$labelComp1 = New-Object System.Windows.Forms.Label
$labelComp1.Location = New-Object System.Drawing.Size(120,5)
$labelComp1.Size = New-Object System.Drawing.Size(450,30)
$labelComp1.Text = "Missing collections for " + $computerName
$labelComp1.Font = [System.Drawing.Font]::new('Segoe UI',13,[System.Drawing.FontStyle]::Bold)
$labelComp1.ForeColor = [System.Drawing.Color]::Black
$labelComp1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelComp1)

$labelIP = New-Object System.Windows.Forms.Label
$labelIP.Location = New-Object System.Drawing.Size(200,40)
$labelIP.Size = New-Object System.Drawing.Size(110,25)
$labelIP.Text = 'Detected IP: '
$labelIP.Font = [System.Drawing.Font]::new('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
$labelIP.ForeColor = [System.Drawing.Color]::Black
$labelIP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelIP)

$IPaddr = New-Object System.Windows.Forms.Label
$IPaddr.Location = New-Object System.Drawing.Size(310,40)
$IPaddr.Size = New-Object System.Drawing.Size(200,25)
$IPaddr.ForeColor = 'Green'
$IPaddr.Text = ''
$IPaddr.Font = [System.Drawing.Font]::new('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
$IPaddr.ForeColor = [System.Drawing.Color]::Black
$IPaddr.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($IPaddr)

$labelSelect = New-Object System.Windows.Forms.Label
$labelSelect.Location = New-Object System.Drawing.Size(10,90)
$labelSelect.Size = New-Object System.Drawing.Size(300,25)
$labelSelect.Text = 'Select collections to add device to:'
$labelSelect.Font = [System.Drawing.Font]::new('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
$labelSelect.ForeColor = [System.Drawing.Color]::Black
$labelSelect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelSelect)

$boxCollection1 = New-Object System.Windows.Forms.CheckedListBox
$boxCollection1.Location = New-Object System.Drawing.Point(10,120)
$boxCollection1.Size = New-Object System.Drawing.Size(560,400)
#$boxCollection1.AutoSize = $false
$boxCollection1.CheckOnClick = $true
$boxCollection1.Font = ‘Segoe UI,11’
$boxCollection1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxCollection1)

$fix = {
	$labelResults.Text = "Adding device to selected collections..."

	foreach ($item in $boxCollection1.CheckedItems) {
		Add-CMDeviceCollectionDirectMembershipRule -CollectionName $item -ResourceId $computer.ResourceID
	}

	$labelResults.Text = "Done! Please WAIT 30 SECONDS, close this window and click COMPARE again."
}

$buttonFix = New-Object System.Windows.Forms.Button
$buttonFix.Location = New-Object System.Drawing.Point(170,545)
$buttonFix.Size = New-Object System.Drawing.Size(250,30)
$buttonFix.Text = 'ADD TO SELECTED COLLECTIONS'
$buttonFix.Font = ‘Segoe UI,11’
$buttonFix.ForeColor = [System.Drawing.Color]::White
$buttonFix.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonFix.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonFix.BackColor = [System.Drawing.Color]::FromArgb(91,181,91)
$buttonFix.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkGreen
$buttonFix.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::Green
$form.Controls.Add($buttonFix)
$buttonFix.Add_Click($fix) #Adds function to button

$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Location = New-Object System.Drawing.Point(10,600)
$labelResults.Size = New-Object System.Drawing.Size(600,30)
$labelResults.Text = ''
$labelResults.Font = [System.Drawing.Font]::new('Segoe UI',10,[System.Drawing.FontStyle]::Bold+[System.Drawing.FontStyle]::Italic)
$labelResults.ForeColor = [System.Drawing.Color]::Black
$labelResults.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelResults)

# POPUP MESSAGE WHEN USER CLOSES APP
$form.Add_Closing({param($sender,$e)
    $result = [System.Windows.Forms.MessageBox]::Show(`
        'Do you want to close this menu?', `
        'Close', [System.Windows.Forms.MessageBoxButtons]::YesNo)
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

# ADD VALUE TO LISTS
if ([string]::IsNullOrEmpty($collectionArray)) {
    $labelResults.Text = "Unable to retrieve data, please close all windows and try again."
} else {
    $labelResults.Text = "Successfully retrieved missing collections for $computerName."
}

foreach ($item in $collectionArray) {
	$boxCollection1.Items.Add($item,$false) | Out-Null
}

# CONNECTS TO SCCM
connectSCCM

# GET DEVICE INFO
$computer = getDevice -hostname $computerName
$IPaddr.ForeColor = 'Green'
$IPaddr.Text = $computer.IPv4

# SHOW FORM
$Null = $form.ShowDialog()