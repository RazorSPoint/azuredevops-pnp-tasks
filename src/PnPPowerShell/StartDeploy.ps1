[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {
    Import-VstsLocStrings "$PSScriptRoot/task.json"
	
    . "$PSScriptRoot/ps_modules/CommonScripts/Utility.ps1"

    [string]$SharePointVersion = Get-VstsInput -Name SharePointVersion

    [string]$FileOrInline = Get-VstsInput -Name FileOrInline

    [string]$PnPPsFilePath = ""

	$FileOrInline = Get-VstsInput -Name 'FileOrInline'
    if ("$FileOrInline".ToUpperInvariant() -eq "FILEPATH") {
        $input_filePath = Get-VstsInput -Name 'PnPPowerShellFilePath' -Require

        try {
            Assert-VstsPath -LiteralPath $input_filePath -PathType Leaf
        } catch {
            Write-Error (Get-VstsLocString -Key 'PS_InvalidFilePath' -ArgumentList $input_filePath)
        }

        if (!$input_filePath.ToUpperInvariant().EndsWith('.PS1')) {
            Write-Error (Get-VstsLocString -Key 'PS_InvalidFilePath' -ArgumentList $input_filePath)
        }

        #$input_arguments = Get-VstsInput -Name 'arguments'
    } else {
        $input_script = Get-VstsInput -Name 'PnPPowerShellInline'
    }

    $agentToolsPath = Get-VstsTaskVariable -Name 'agent.toolsDirectory' -Require
    $modulePath = Get-PnPPackageModulePath -SharePointVersion $SharePointVersion -AgentToolPath $agentToolsPath
    Load-PnPPackages -SharePointVersion $SharePointVersion -AgentToolPath $agentToolsPath
 
	#Write-Host (Get-VstsLocString -Key 'GeneratingScript')
    $contents = @()
    # $contents += "`$ErrorActionPreference = '$input_errorActionPreference'"

    $contents += "`$null = Import-Module $modulePath -DisableNameChecking -Verbose:`$false"
    if ("$input_targetType".ToUpperInvariant() -eq 'FILE') {
        $contents += ". '$("$input_filePath".Replace("'", "''"))' $input_arguments".Trim()
        #Write-Host (Get-VstsLocString -Key 'PS_FormattedCommand' -ArgumentList ($contents[-1]))
    } else {
        $contents += "$input_script".Replace("`r`n", "`n").Replace("`n", "`r`n")
    }

	# Write the script to disk.
    #Assert-VstsAgent -Minimum '2.115.0'
    $tempDirectory = Get-VstsTaskVariable -Name 'agent.tempDirectory' -Require
    Assert-VstsPath -LiteralPath $tempDirectory -PathType 'Container'
    $filePath = [System.IO.Path]::Combine($tempDirectory, "$([System.Guid]::NewGuid()).ps1")
    $joinedContents = [System.String]::Join(([System.Environment]::NewLine), $contents)
    $null = [System.IO.File]::WriteAllText($filePath,$joinedContents,([System.Text.Encoding]::UTF8))

	# Prepare the external command values.
    #
    # Note, use "-Command" instead of "-File". On PowerShell v4 and V3 when using "-File", terminating
    # errors do not cause a non-zero exit code.
    $powershellPath = Get-Command -Name powershell.exe -CommandType Application | Select-Object -First 1 -ExpandProperty Path
    Assert-VstsPath -LiteralPath $powershellPath -PathType 'Leaf'
    $arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Unrestricted -Command `". '$($filePath.Replace("'", "''"))'`""
    $splat = @{
        'FileName' = $powershellPath
        'Arguments' = $arguments
        'WorkingDirectory' = $input_workingDirectory
    }	

	# Switch to "Continue".
    $global:ErrorActionPreference = 'Continue'
    $failed = $false

	Invoke-VstsTool @splat
	
}
catch {
    $ErrorMessage = $_.Exception.Message
    throw "An Error occured. The error message was: $ErrorMessage. `n Stackstace `n $($_.ScriptStackTrace)"
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation  
}
    
