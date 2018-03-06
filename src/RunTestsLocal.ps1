cls
#Invoke-Pester './Tests/*' -EnableExit

. "DeployArtifactsWithPnP\Scripts\Utility.ps1"
$agentToolPath = "C:\temp"
Load-PnPPackages -SharePointVersion "SpOnline" -AgentToolPath $agentToolPath

Exit

