function Install-ZipFolderResource {
    param
    (
        [string]$ZipPath,
        [string]$ZipFileName,    
        [string]$Out
    )

    #is used to unpack the zips if they are not extracted already
    #unpacked DLLs are 10 times higher than packed.
    #this saves up- and download time.

    $zipFolderName = [io.path]::GetFileNameWithoutExtension($ZipFileName)

    if ((Test-Path "$ZipPath/$zipFolderName/") -eq $false) {

        Add-Type -AssemblyName System.IO.Compression.FileSystem

        [System.IO.Compression.ZipFile]::ExtractToDirectory("$ZipPath/$ZipFileName", $Out)
    }
}

function Copy-FileRemote {
    param(
        [string]$adminusername,
        [string]$adminpassword,
        [string]$copysource,
        [string]$copytarget,
        [string]$machine,
        [string]$cleantarget = 'false'
    ) 

    Write-Verbose "AdminUserName = $adminusername" 
    Write-Verbose "CopySource = $copysource"
    Write-Verbose "CopyTarget = $copytarget"
    Write-Verbose "Machine = $machine"
 
    $b_clean = [System.Convert]::ToBoolean($cleantarget)

    try {
        $user = $adminusername
        $password = $adminpassword
        $encryptedpassword = Convertto-SecureString -AsPlainText -Force -String $password
        $cred = New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $user, $encryptedpassword
    
        Write-Verbose "Opening new session on machine $machine"
        $session = New-PSSession -ComputerName $machine -Credential $cred

        ### START Coppy
        Write-Verbose "Copy-Item process on $machine begins"
        Copy-Item -Path $copysource -Destination $copytarget -ToSession $session -Force 
        Write-Verbose "Copy-Item process on $machine finished"
   
    }
    catch
	{
		$ErrorMessage = $_.Exception.Message
		throw "An Error occured. The error message was: $ErrorMessage"
	}
    finally {
        Remove-PSSession -Session $session
        #extra actions
    }
}

function Remove-FileRemote {
    param(
        [string]$adminusername,
        [string]$adminpassword,
        [string]$targetFile,
        [string]$machine
    ) 

    Write-Verbose "AdminUserName = $adminusername" 
    Write-Verbose "TargetFile = $targetFile"
    Write-Verbose "Machine = $machine"

    try {
        $user = $adminusername
        $password = $adminpassword
        $encryptedpassword = Convertto-SecureString -AsPlainText -Force -String $password
        $cred = New-Object -Typename System.Management.Automation.PSCredential -Argumentlist $user, $encryptedpassword
    
        Write-Verbose "Deleting file $targetFile on machine $machine"
        $session = New-PSSession -ComputerName $machine -Credential $cred

        Invoke-Command -session $session -Scriptblock {
            Remove-Item -Path $args[0]
        } -ArgumentList $targetFile        
          
    }
    catch
	{
		$ErrorMessage = $_.Exception.Message
		throw "An Error occured. The error message was: $ErrorMessage"
	}
    finally {
        #extra actions
        Remove-PSSession -Session $session
    }
}

function Invoke-VstsRemotePowerShellJob {
    param (
        [string]$fqdn, 
        [string]$scriptToInvoke,
        [string]$adminusername,
        [string]$adminpassword,
		[string]$tempScriptPath = "C:\temp",
        [string]$scriptArguments = "",
        [string]$port = "5985",
        [string]$initializationScriptPath = "",
        [string]$httpProtocolOption = "-UseHttp",
        [string]$skipCACheckOption = "",
        [string]$enableDetailedLogging = "false",
        [string]$sessionVariables = ""
    )

    Write-Verbose "fqdn = $fqdn"
	Write-Verbose "tempScriptPath = $tempScriptPath"
    Write-Verbose "port = $port"
    Write-Verbose "scriptArguments = $scriptArguments"
    Write-Verbose "initializationScriptPath = $initializationScriptPath"
    Write-Verbose "protocolOption = $httpProtocolOption"
    Write-Verbose "skipCACheckOption = $skipCACheckOption"
    Write-Verbose "enableDetailedLogging = $enableDetailedLogging"

    # check if we are on a hosted or a private agent
    if (Test-Path "$env:AGENT_HOMEDIRECTORY\Agent\Worker") {
        #is hosted
        Get-ChildItem $env:AGENT_HOMEDIRECTORY\Agent\Worker\*.dll | ForEach-Object {
            [void][reflection.assembly]::LoadFrom( $_.FullName )
            Write-Verbose "Loading .NET assembly:`t$($_.name)"
        }

        Get-ChildItem $env:AGENT_HOMEDIRECTORY\Agent\Worker\Modules\Microsoft.TeamFoundation.DistributedTask.Task.DevTestLabs\*.dll | % {
            [void][reflection.assembly]::LoadFrom( $_.FullName )
            Write-Verbose "Loading .NET assembly:`t$($_.name)"
        }
    }
    else {
        #is private
        if (Test-Path "$env:AGENT_HOMEDIRECTORY\externals\vstshost") {        
            Import-Module "$env:AGENT_HOMEDIRECTORY\externals\vstshost\Microsoft.TeamFoundation.DistributedTask.Task.LegacySDK.dll"
            Write-Verbose "Loading .NET assembly: Microsoft.TeamFoundation.DistributedTask.Task.LegacySDK.dll"
        }
    }

    #enable verbose logging
    $enableDetailedLoggingOption = ''
    if ($enableDetailedLogging -eq "true") {
        $enableDetailedLoggingOption = '-EnableDetailedLogging'
    }

    $parsedSessionVariables = Get-ParsedSessionVariables -inputSessionVariables $sessionVariables
   
    Write-Verbose "Creating temporary script file"

    #create temporary ps file
    $tmpFileName = [System.IO.Path]::GetRandomFileName()    
    $filePath = "$tempScriptPath\$tmpFileName.ps1"	
    $scriptToInvoke | Out-File -FilePath $filePath -Append

    #copy file to remote machine
    Copy-FileRemote -adminusername $adminusername -adminpassword $adminpassword -copysource $filePath -copytarget $tempScriptPath -machine $fqdn

    Write-Verbose "Initiating deployment on $fqdn"

    $deploymentResponse = ""

    #execute remote copied file
    try {    

        $credential = new-object System.Net.NetworkCredential($adminusername, $adminpassword);
        #prepare invoke command of "Invoke-PsOnRemote"
		
        [String]$psOnRemoteScriptBlockString = "Invoke-PsOnRemote -MachineDnsName $fqdn -ScriptPath `$filePath -WinRMPort $port -Credential `$credential -ScriptArguments `$scriptArguments -InitializationScriptPath `$initializationScriptPath -SessionVariables `$parsedSessionVariables $skipCACheckOption $httpProtocolOption $enableDetailedLoggingOption"
		Write-Verbose $psOnRemoteScriptBlockString
		
        [scriptblock]$psOnRemoteScriptBlock = [scriptblock]::Create($psOnRemoteScriptBlockString)
        $deploymentResponse = Invoke-Command -ScriptBlock $psOnRemoteScriptBlock 
        
        Write-Verbose $deploymentResponse   
    }
    catch
	{
		$ErrorMessage = $_.Exception.Message
		throw "An Error occured. The error message was: $ErrorMessage"
	}
    finally {
        #remove remote file
        Remove-FileRemote -adminusername $adminusername -adminpassword $adminpassword -targetFile $filePath -machine $fqdn
        Remove-Item -Path $filePath -Force -ErrorAction SilentlyContinue

        $status = $deploymentResponse.Status

        if ($deploymentResponse -ne "" -and $status -ne "Passed") {
            Write-Verbose $deploymentResponse.Error.ToString()
            $errorMessage = $deploymentResponse.Error.Message
            throw $errorMessage
        }
    }   
}
