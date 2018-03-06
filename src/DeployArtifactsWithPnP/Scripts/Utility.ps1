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

<#
.Synopsis
    
.DESCRIPTION
    
.EXAMPLE
   
#>
function Load-PnPPackages {

	[CmdletBinding()]
    param(
		[Parameter(Mandatory=$true, Position=0)]   
        [string]$SharePointVersion,
        [Parameter(Mandatory=$true, Position=1)]   
        [string]$AgentToolPath        
	)

    $pnpModuleName = ""
    $pnpDllName = ""
    # suppress output
		switch ($SharePointVersion){
        "Sp2013" {
            $pnpModuleName = "SharePointPnPPowerShell2013"  
            $pnpDllName = "SharePointPnP.PowerShell.2013.Commands.dll"
        }
		"Sp2016" { 
            $pnpModuleName = "SharePointPnPPowerShell2016"
            $pnpDllName = "SharePointPnP.PowerShell.2016.Commands.dll"
		}
		"SpOnline" {
            $pnpModuleName = "SharePointPnPPowerShellOnline"
            $pnpDllName = "SharePointPnP.PowerShell.Online.Commands.dll"
        }
		default { throw "Only SharePoint 2013, 2016 or SharePoint Online are supported at the moment" }
	}
     
 <#   try{#>
        #check for PSGallery entry and add if not present
        $psRepositoriy = Get-PSRepository -Name "PSGallery"

        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        
        if ($psRepositoriy -eq $null) {
            Register-PSRepository -Default                        
        }

        
        $pnpModule = Find-Module -Name $pnpModuleName
        $modulePath = "$AgentToolPath\$pnpModuleName\$($pnpModule.Version)\$pnpDllName"

        if(-not (Test-Path -Path $modulePath)){
            Save-Module -Name $pnpModuleName -Path $AgentToolPath
        }else{
            Write-Host "Module $pnpModuleName with version $($pnpModule.Version) is already downloaded." -ForegroundColor Yellow
        }
    
        
        Import-Module $modulePath -DisableNameChecking -Verbose:$false
    
        Write-Host "Assemblies loaded." -ForegroundColor Green

  <#  }catch{

        $ErrorMessage = $_.Exception.Message
        Write-Host $ErrorMessage -ForegroundColor Red

        return $false
    }   #>

    return $true
}

