[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {

	if(-not($PSVersionTable) -or -not($PSVersionTable.PSVersion.Major -gt 5)){
		Throw "PowerShell 5 is not installed on the agent."
	}

	Import-VstsLocStrings "$PSScriptRoot/task.json"
	
	. "$PSScriptRoot/Scripts/PnPAppHelper.ps1"
	. "$PSScriptRoot/Scripts/Utility.ps1"

	Install-ZipFolderResource -ZipPath "$PSScriptRoot/ps_modules" -ZipFileName "PnP.zip" -Out ".\ps_modules\"

    [string]$SharePointVersion = Get-VstsInput -Name SharePointVersion
	
	[string]$AppFilePath = Get-VstsInput -Name AppFilePath
	if (-not (Test-Path $AppFilePath)) 
	{
		Throw "File path '$AppFilePath' for variable `$AppFilePath does not exist."
	}
	

    [string]$WebUrl = Get-VstsInput -Name TargetWebUrl
	if(($WebUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false){
		Throw "web url '$WebUrl' of the variable `$WebUrl is not a valid url. E.g. http://my.sharepoint.sitecollection."
	}


	[string]$AppVersion = Get-VstsInput -Name AppVersion
	if(-not([string]::IsNullOrEmpty($AppVersion)) -and ($AppVersion -match "(?:(\d+)\.)?(?:(\d+)\.)?(?:(\d+)\.\d+)") -eq $false){
		Throw "version number '$AppVersion' of variable `$AppVersion is not a valid .NET version number. Correct pattern is Mayor.Minor.Revision.Buidlnumber (e.g. 1.0.1.12)."
	}


    [string]$DeployUserName = Get-VstsInput -Name AdminLogin

    [string]$DeployPassword = Get-VstsInput -Name AdmninPassword

    [string]$RemoteMachine = Get-VstsInput -Name RemoteMachine

	$hasMachineConnection = Test-Connection -ComputerName $RemoteMachine -Quiet -Count 4

	if($hasMachineConnection -eq $false){
		Throw "the remote machine '$RemoteMachine' of variable `$RemoteMachine is not responding. Make sure the name is correct (try to use the FQDN) and the machine is connected."
	}else{
		 $RemoteMachine = ([System.Net.Dns]::GetHostByName($RemoteMachine)).HostName
	}

	[bool]$AddInType = Get-VstsInput -Name AddInType -AsBool
	
	[bool]$ReinstallIfExisting = Get-VstsInput -Name ReinstallIfExisting -AsBool

	if($AddInType -eq $true){

		[string]$ClientId = Get-VstsInput -Name ClientId
		if(-not([string]::IsNullOrEmpty($ClientId)) -and -not($ClientId -match "^[{(]?[0-9A-F]{8}[-]?([0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$")){
			Throw "the client id '$ClientId' of variable `$ClientId is not a valid guid."
		}

		[string]$RemoteAppUrl = Get-VstsInput -Name RemoteAppUrl
		if(($RemoteAppUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false){
			Throw "web url '$RemoteAppUrl' of variable `$RemoteAppUrl is not a valid url. E.g. http://my.sharepoint.sitecollection."
		}

		Write-Host "Calling ."".\Scripts\DeployProviderHostedApp.ps1"" -AppFilePath $AppFilePath -WebUrl $WebUrl -RemoteAppUrl $RemoteAppUrl -DeployUserName $DeployUserName -DeployPassword $DeployPassword  -SharePointVersion $SharePointVersion -Machine $RemoteMachine  -AppVersion $AppVersion -ReinstallIfExisting:$ReinstallIfExisting -ClientId $ClientId -Verbose"
		.".\Scripts\DeployProviderHostedApp.ps1" -AppFilePath $AppFilePath -WebUrl $WebUrl -RemoteAppUrl $RemoteAppUrl -DeployUserName $DeployUserName -DeployPassword $DeployPassword  -SharePointVersion $SharePointVersion -Machine $RemoteMachine -AppVersion $AppVersion -ReinstallIfExisting:$ReinstallIfExisting -ClientId $ClientId -Verbose

	}else{
		Write-Host "Calling ."".\Scripts\DeploySharePointHostedApp.ps1"" -AppFilePath $AppFilePath -WebUrl $WebUrl -DeployUserName $DeployUserName -DeployPassword $DeployPassword -SharePointVersion $SharePointVersion -Machine $RemoteMachine -AppVersion $AppVersion -ReinstallIfExisting:$ReinstallIfExisting -Verbose"
		.".\Scripts\DeploySharePointHostedApp.ps1" -AppFilePath $AppFilePath -WebUrl $WebUrl -DeployUserName $DeployUserName -DeployPassword $DeployPassword -SharePointVersion $SharePointVersion -Machine $RemoteMachine -AppVersion $AppVersion -ReinstallIfExisting:$ReinstallIfExisting -Verbose
	}
	
} finally {
    Trace-VstsLeavingInvocation $MyInvocation
}
