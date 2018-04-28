Import-Module .\ps_modules\VstsTaskSdk

# Input 'MyInput':
$env:INPUT_FileOrInline = "File"
$env:INPUT_PnPPowerShellInline = "Write-Host 'Test!!'"
$env:AGENT_TEMPDIRECTORY = "C:\temp\"

Invoke-VstsTaskScript -ScriptBlock { . .\..\src\PnPPowerShell\StartDeploy.ps1 }

Exit