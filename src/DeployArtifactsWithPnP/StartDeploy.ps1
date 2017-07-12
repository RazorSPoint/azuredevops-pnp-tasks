[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {
	Import-VstsLocStrings "$PSScriptRoot/task.json"
	
	. "$PSScriptRoot/Scripts/PnPAppHelper.ps1"

	Install-ZipFolderResource -ZipPath "$PSScriptRoot/ps_modules" -ZipFileName "PnP.zip" -Out ".\ps_modules\"

    [string]$SharePointVersion = Get-VstsInput -Name SharePointVersion
	
	[string]$PnPXmlFilePath = Get-VstsInput -Name PnPXmlFilePath
	if (-not (Test-Path $PnPXmlFilePath)) 
	{
		Throw "File path '$PnPXmlFilePath' for variable `$PnPXmlFilePath does not exist."
	}
	
    [string]$WebUrl = Get-VstsInput -Name TargetWebUrl
	if(($WebUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false){
		Throw "web url '$WebUrl' of the variable `$WebUrl is not a valid url. E.g. http://my.sharepoint.sitecollection."
	}

	[string]$Handlers = (Get-VstsInput -Name Handlers)

	[string]$Parameters = (Get-VstsInput -Name Parameters)
	
    [string]$DeployUserName = Get-VstsInput -Name AdminLogin

    [string]$DeployPassword = Get-VstsInput -Name AdmninPassword

	#preparing pnp provisioning
	Load-Assemblies $SharePointVersion

	$secpasswd = ConvertTo-SecureString $DeployPassword -AsPlainText -Force
	$adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, $secpasswd)

	Write-Host "Connect to '$WebUrl' as '$DeployUserName'..."
	Connect-PnPOnline -Url $WebUrl -Credentials $adminCredentials
	Write-Host "Successfully connected to '$WebUrl'..." 
	

	$ProvParams = @{ 
		Path = $PnPXmlFilePath
	} 

	#check for handlers
	if(-not [string]::IsNullOrEmpty($Handlers)){
		$ProvParams.Handlers = $Handlers.split(",;").join(",")
	}

	#check for parameters
	if(-not [string]::IsNullOrEmpty($Parameters)){
		$ProvParams.Parameters = $Parameters
	}

	#execute provisioning
	Apply-PnPProvisioningTemplate @ProvParams

}catch {
		$ErrorMessage = $_.Exception.Message
		throw "An Error occured. The error message was: $ErrorMessage"
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation
}
