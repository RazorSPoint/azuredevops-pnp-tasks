[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {
	Import-VstsLocStrings "$PSScriptRoot/task.json"
	
	#. "$PSScriptRoot/Scripts/PnPAppHelper.ps1"

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

	[bool]$UseSpecificHandlers = Get-VstsInput -Name UseSpecificHandlers -AsBool
	
} finally {
    Trace-VstsLeavingInvocation $MyInvocation
}
