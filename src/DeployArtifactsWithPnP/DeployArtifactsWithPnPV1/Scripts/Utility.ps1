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
function Load-Assemblies {

	[CmdletBinding()]
    param(
		[Parameter(Mandatory=$true, Position=0)]   
		[ValidateSet('SpOnline','Sp2016','Sp2013')]
		[string]$SharePointVersion
	)

    # suppress output
		switch ($SharePointVersion){
		"Sp2016" { 
			$modulePath = '.\ps_modules\PnP\SharePointPnPPowerShell2016\SharePointPnP.PowerShell.2016.Commands.dll'			
		}
		"SpOnline" {$modulePath = '.\ps_modules\PnP\SharePointPnPPowerShellOnline\SharePointPnP.PowerShell.Online.Commands.dll'}
		default { throw "Only SharEPoint 2016 or SharePoint Online are supported at the moment" }
	}
     
	Import-Module $modulePath -DisableNameChecking -Verbose:$false

    Write-Output "Assemblies loaded."
}