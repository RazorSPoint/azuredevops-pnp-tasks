[CmdletBinding()]
param()

# For more information on the VSTS Task SDK:
# https://github.com/Microsoft/vsts-task-lib

Trace-VstsEnteringInvocation $MyInvocation

try {	
    
    Write-Host "The script was partially inspired by the official VSTS inline PowerShell task from Microsoft Corporation. Some lines are reused.
Source Code can be found here https://github.com/Microsoft/vsts-tasks/tree/e9f6da2c523e456f10421ed40dbeed1dd45af2b4/Tasks/powerShell"

    . "$PSScriptRoot/ps_modules/CommonScripts/Utility.ps1"
	
	###############
	#Get inputs
	###############
    [string]$input_SharePointVersion = Get-VstsInput -Name SharePointVersion
    [string]$input_FileOrInline = Get-VstsInput -Name FileOrInline

    $ConnectedService = Get-VstsInput -Name ConnectedServiceName -Require
    $ServiceEndpoint = (Get-VstsEndpoint -Name $ConnectedService -Require)

    [string]$WebUrl = $ServiceEndpoint.Url
    if (($WebUrl -match "(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)") -eq $false) {
       #Write-VstsTaskError -Message "`nweb url '$WebUrl' of the variable `$WebUrl is not a valid url. E.g. http://my.sharepoint.sitecollection.`n"
    }

    [string]$DeployUserName = $ServiceEndpoint.Auth.parameters.username
    [string]$DeployPassword = $ServiceEndpoint.Auth.parameters.password
    
    $input_ErrorActionPreference = Get-VstsInput -Name 'errorActionPreference' -Default 'Stop'
    switch ($input_ErrorActionPreference.ToUpperInvariant()) {
        'STOP' { }
        'CONTINUE' { }
        'SILENTLYCONTINUE' { }
        default {
            Write-VstsError -Message "Invalid ErrorActionPreference '$input_ErrorActionPreference'. The value must be one of: 'Stop', 'Continue', or 'SilentlyContinue'"
        }
    }
    [bool]$input_FailOnStderr = Get-VstsInput -Name 'failOnStderr' -AsBool
    
	
    $input_WorkingDirectory = Get-VstsInput -Name 'workingDirectory' -Require
    Assert-VstsPath -LiteralPath $input_WorkingDirectory -PathType 'Container'

	$input_FileOrInline = Get-VstsInput -Name 'FileOrInline'
    if ("$input_FileOrInline".ToUpperInvariant() -eq "FILEPATH") {
        $psfilePath = Get-VstsInput -Name 'PnPPowerShellFilePath' -Require

        try {
            Assert-VstsPath -LiteralPath $psfilePath -PathType Leaf
        } catch {
            Write-Error "Path to the PowerShell file $psfilePath does not exist"
        }

        if (!$psfilePath.ToUpperInvariant().EndsWith('.PS1')) {
            Write-Error "The given file $psfilePath is not a PowerShell script."
        }

        $PsArguments = Get-VstsInput -Name 'PsArguments'
    } else {
        $psInlineScript = Get-VstsInput -Name 'PnPPowerShellInline'
    }

    ########################
    # Load the PnP Modules
    ########################
    $agentToolsPath = Get-VstsTaskVariable -Name 'agent.toolsDirectory' -Require
    $modulePath = Get-PnPPackageModulePath -SharePointVersion $input_SharePointVersion -AgentToolPath $agentToolsPath
    $null = Load-PnPPackages -SharePointVersion $input_SharePointVersion -AgentToolPath $agentToolsPath
 
	#############
	# generate the script
	#############
	Write-Host "Generating the scripts..."
    $contents = @()
    $contents += "`$ErrorActionPreference = '$input_ErrorActionPreference'"

	#add the ps line to include the module into the script
    $contents += "`$null = Import-Module $modulePath -DisableNameChecking -Verbose:`$false"
    
    #connect to SharePoint online
    $contents += "`$null = Import-Module $modulePath -DisableNameChecking -Verbose:`$false"

    $contents += "`$secpasswd = ConvertTo-SecureString '$DeployPassword' -AsPlainText -Force"
    $contents += "`$adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, `$secpasswd)"

    $contents += "Write-Host `"`nConnect to '$WebUrl' as '$DeployUserName'...`""
    $contents += "Connect-PnPOnline -Url '$WebUrl' -Credentials `$adminCredentials"
    $contents += "Write-Host `"Successfully connected to '$WebUrl'...`n`"" 

    if ("$input_targetType".ToUpperInvariant() -eq 'FILE') {
        $contents += ". '$("$psfilePath".Replace("'", "''"))' $input_arguments".Trim()
        Write-Host "Formatted command: $($contents[-1])"
    } else {
        $contents += "$psInlineScript".Replace("`r`n", "`n").Replace("`n", "`r`n")
    }

	#############
	# save the script to temp folder.
	#############
    $tempDirectory = Get-VstsTaskVariable -Name 'agent.tempDirectory' -Require
    Assert-VstsPath -LiteralPath $tempDirectory -PathType 'Container'
    $filePath = [System.IO.Path]::Combine($tempDirectory, "$([System.Guid]::NewGuid()).ps1")
    $joinedContents = [System.String]::Join(([System.Environment]::NewLine), $contents)
    $null = [System.IO.File]::WriteAllText($filePath,$joinedContents,([System.Text.Encoding]::UTF8))

	# create powershell call with the script.
    $powershellPath = Get-Command -Name powershell.exe -CommandType Application | Select-Object -First 1 -ExpandProperty Path
    Assert-VstsPath -LiteralPath $powershellPath -PathType 'Leaf'
    $arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Unrestricted -Command `". '$($filePath.Replace("'", "''"))'`""
    $splat = @{
        'FileName' = $powershellPath
        'Arguments' = $arguments
        'WorkingDirectory' = $input_WorkingDirectory
    }
	
	###############
	# Switch to "Continue".
	###############
    $global:ErrorActionPreference = 'Continue'

	###############
	# Run the script.
	###############
    if (!$input_FailOnStderr) {
        Invoke-VstsTool @splat
    } else {
        $inError = $false
        $errorLines = New-Object System.Text.StringBuilder
        Invoke-VstsTool @splat 2>&1 |
            ForEach-Object {
                if ($_ -is [System.Management.Automation.ErrorRecord]) {
                    # Buffer the error lines.
                    $failed = $true
                    $inError = $true
                    $null = $errorLines.AppendLine("$($_.Exception.Message)")

                    # Write to verbose to mitigate if the process hangs.
                    Write-Verbose "STDERR: $($_.Exception.Message)"
                } else {
                    # Flush the error buffer.
                    if ($inError) {
                        $inError = $false
                        $message = $errorLines.ToString().Trim()
                        $null = $errorLines.Clear()
                        if ($message) {
                            Write-VstsTaskError -Message $message
                        }
                    }

                    Write-Host "$_"
                }
            }

        # Flush the error buffer one last time.
        if ($inError) {
            $inError = $false
            $message = $errorLines.ToString().Trim()
            $null = $errorLines.Clear()
            if ($message) {
                Write-VstsTaskError -Message $message
            }
        }
    }

    # Fail on $LASTEXITCODE
    if (!(Test-Path -LiteralPath 'variable:\LASTEXITCODE')) {
        $failed = $true
        Write-Verbose "Unable to determine exit code"
        Write-VstsTaskError -Message "Unexpected exception. Unable to determine the exit code from powershell."
    } else {
        if ($LASTEXITCODE -ne 0) {
            $failed = $true
            Write-VstsTaskError -Message "PowerShell exited with code: $LASTEXITCODE"
        }
    }

    # Fail if any errors.
    if ($failed) {
        Write-VstsSetResult -Result 'Failed' -Message "Error detected" -DoNotThrow
    }	
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation  
}
    
