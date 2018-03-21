[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {
    Import-VstsLocStrings "$PSScriptRoot/task.json"
	
    . "$PSScriptRoot/Scripts/Utility.ps1"
    # get the tmp path of the agent
    $agentTmpPath = "$($env:AGENT_RELEASEDIRECTORY)\_temp"
    $tmpInlineXmlFileName = [System.IO.Path]::GetRandomFileName() + ".xml"

    [string]$SharePointVersion = Get-VstsInput -Name SharePointVersion
		
    [string]$WebUrl = Get-VstsInput -Name TargetWebUrl
    if (($WebUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false) {
        Throw "web url '$WebUrl' of the variable `$WebUrl is not a valid url. E.g. http://my.sharepoint.sitecollection."
    }

    [string]$FileOrInline = Get-VstsInput -Name FileOrInline

    [string]$PnPXmlFilePath = ""

    if ($FileOrInline -eq "File") {
        [string]$PnPXmlFilePath = Get-VstsInput -Name PnPXmlFilePath
        if (-not (Test-Path $PnPXmlFilePath)) {
            Throw "File path '$PnPXmlFilePath' for variable `$PnPXmlFilePath does not exist."
        }
    }
    else {

        #get xml string and check for valid xml
        [string]$PnPXmlInline = (Get-VstsInput -Name PnPXmlInline)
		
        $PnPXml = New-Object System.Xml.XmlDocument
        try {
            $PnPXmlFilePath = "$agentTmpPath/$tmpInlineXmlFileName"
            #if patrh not exists, create it!
            if (-not (Test-Path -Path $agentTmpPath)) {
                New-Item -ItemType Directory -Force -Path $agentTmpPath
            }
            $PnPXml.LoadXml($PnPXmlInline)
            $PnPXml.Save($PnPXmlFilePath)
        }
        catch [System.Xml.XmlException] {
            throw "$($_.toString())"		
        }
    }

    [string]$Handlers = (Get-VstsInput -Name Handlers)

    $TmpParameters = (Get-VstsInput -Name Parameters)
	
    [string]$DeployUserName = Get-VstsInput -Name AdminLogin

    [string]$DeployPassword = Get-VstsInput -Name AdmninPassword
	
    [bool]$ClearNavigation = Get-VstsInput -Name ClearNavigation -AsBool

    [bool]$IgnoreDuplicateDataRowErrors = Get-VstsInput -Name IgnoreDuplicateDataRowErrors -AsBool

    [bool]$OverwriteSystemPropertyBagValues = Get-VstsInput -Name OverwriteSystemPropertyBagValues -AsBool

    [bool]$ProvisionContentTypesToSubWebs = Get-VstsInput -Name ProvisionContentTypesToSubWebs -AsBool

    #preparing pnp provisioning
    $agentToolsPath = "$($env:AGENT_WORKFOLDER)\_tool"
    Load-PnPPackages -SharePointVersion $SharePointVersion -AgentToolPath $agentToolsPath

    $secpasswd = ConvertTo-SecureString $DeployPassword -AsPlainText -Force
    $adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, $secpasswd)

    Write-Host "Connect to '$WebUrl' as '$DeployUserName'..."
    Connect-PnPOnline -Url $WebUrl -Credentials $adminCredentials
    Write-Host "Successfully connected to '$WebUrl'..." 
	

    $ProvParams = @{ 
        Path                             = $PnPXmlFilePath
        ClearNavigation                  = $ClearNavigation
        IgnoreDuplicateDataRowErrors     = $IgnoreDuplicateDataRowErrors
        OverwriteSystemPropertyBagValues = $OverwriteSystemPropertyBagValues
        ProvisionContentTypesToSubWebs   = $ProvisionContentTypesToSubWebs
    } 

    #check for handlers
    if (-not [string]::IsNullOrEmpty($Handlers)) {
        $ProvParams.Handlers = [System.String]::Join(",",$Handlers.split(",;"))
    }

    #check for parameters
    if (-not [string]::IsNullOrEmpty($TmpParameters)) {
        [System.Collections.Hashtable] $Parameters = ConvertFrom-StringData -StringData $TmpParameters
        $ProvParams.Parameters = $Parameters
    }

    #execute provisioning
    Apply-PnPProvisioningTemplate @ProvParams

}
catch {
    $ErrorMessage = $_.Exception.Message
    throw "An Error occured. The error message was: $ErrorMessage. `n Stackstace `n $($_.ScriptStackTrace)"
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation

    #clean up tmp path
    if ($FileOrInline -eq 'Inline' -and (Test-Path $agentTmpPath)) {
        Remove-Item $agentTmpPath -Recurse       
    }
}
    
