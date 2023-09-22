#USE THIS SCRIPT TO UPDATE TOOL
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

$computer = $env:COMPUTERNAME
$date = Get-Date -Format "MM_dd_yy"

$packageSourcePath = "\\jmpist01094p01\Data\Scripting\Khoa - Finished Scripts\SCCM Collection Tool - Full Version\Latest Version"
$localPath = "C:\temp\SCCM Collection Tool - Full Version"

If (!(Test-Path -Path $localPath)) {
	New-Item -Path $localPath -ItemType Directory
}

Copy-Item -Path "$packageSourcePath\*" -Destination $localPath -Recurse -Force

Move-Item -Path "$localPath\SCCM Collection Tool.lnk" -Destination "C:\Users\Public\Desktop"
Move-Item -Path "$localPath\UPDATE - SCCM Collection Tool.lnk" -Destination "C:\Users\Public\Desktop"