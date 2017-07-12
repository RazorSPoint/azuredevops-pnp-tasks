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