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

    Write-Host "Assemblies loaded."
}

<#
.Synopsis
    
.DESCRIPTION
    
.EXAMPLE
   
#>
function Enable-SideLoading {

	[CmdletBinding()]
    param(		
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()]    
		[bool]$Enable = $true, 
		
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateSet('SpOnline','Sp2016','Sp2013')]   
		[string]$SharePointVersion, 

        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]$DeployUserName,

        [Parameter(Mandatory=$true, Position=3)]
        [ValidateNotNullOrEmpty()]
        [string]$DeployPassword,       

		[Parameter(Mandatory=$false, Position=4)]
		[string]$Machine = $null,

		[Parameter(Mandatory=$false, Position=5)]
		[switch]$Force
	)

    . "$PSScriptRoot/Utility.ps1"
    
    # this is the side-loading Feature ID..
    $FeatureId = [GUID]("AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D")

	$siteFeatures = Get-PnPFeature -Scope Site

    $feature = $siteFeatures | Where-Object { $_.DefinitionId -eq $FeatureId } | Select-Object -First 1

    if ($feature)
    {
        if ($Enable) 
        { 
            Write-Host "Feature is already activated in this site." 
            return
        } 
        
		#disable sideloading feature if sp 2016
		if($SharePointVersion -eq "Sp2016"){

			if($Machine){
				Write-Host "Disabeling Sideloading on remote Machine: $Machine"

                $invokeString = "Add-PSSnapin -Name Microsoft.SharePoint.PowerShell; Disable-SPFeature -Identity $FeatureId -Url $WebUrl -Confirm:`$false"
                Invoke-VstsRemotePowerShellJob -fqdn $Machine -scriptToInvoke $invokeString -adminusername $DeployUserName -adminpassword $DeployPassword

			}else{
				Write-Host "Disabeling Sideloading on local Machine."
				Add-PSSnapin -Name *sharepoint*
				Disable-SPFeature -Identity $FeatureId -Url $WebUrl -Confirm:$false
			}
				
		}

		#disable sideloading feature with PnP if sp online
		if($SharePointVersion -eq "SpOnline"){
			Disable-PnPFeature -Identity $FeatureId -Scope Site -Force
		}
		Write-Host "Feature '$FeatureId' successfully deactivated.."        
        
    }
    else
    {
		if($Enable -eq $false){
			Write-Host "The feature is not active at this scope."
            return			
		}
 
	    #enable sideloading feature if sp 2016
		if($SharePointVersion -eq "Sp2016"){

			if($Machine){
				Write-Host "Enabling Sideloading on remote Machine: $Machine"

                $invokeString = "Add-PSSnapin -Name Microsoft.SharePoint.PowerShell; Enable-SPFeature -Identity $FeatureId -Url $WebUrl"
                Invoke-VstsRemotePowerShellJob -fqdn $Machine -scriptToInvoke $invokeString -adminusername $DeployUserName -adminpassword $DeployPassword

			}else{
				Write-Host "Enabling Sideloading on local Machine."
				Add-PSSnapin -Name *sharepoint*
				Enable-SPFeature -Identity $FeatureId -Url $WebUrl
			}

		}

		#enable sideloading feature with PnP if sp online
		if($SharePointVersion -eq "SpOnline"){
			Enable-PnPFeature -Identity $FeatureId -Scope Site
		}

		Write-Host "Feature '$FeatureId' successfully activated.."
   
    }
}

<#
.Synopsis
   Reads the app manifest data from the app package. 
.DESCRIPTION
   Reads the app manifest data from the app package. The packages in unzipped and the app manifest is loaded. 
   The following data is read: client ID, app version, AllowAppPolicy, product ID
.EXAMPLE
   Get-SpAddinPackageInformations -Path "C:work\path\to\file.app"
#>
function Get-SpAddinPackageInformations {

    [CmdletBinding()]
    [OutputType([Hashtable])]
    param(		
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()]  
		[string]$Path
	)

    # Open zip
    Add-Type -assembly  System.IO.Compression.FileSystem
    Write-Host "Open zip file '$Path'..."
    $zip =  [System.IO.Compression.ZipFile]::Open($Path, "Update")

    try{
        $fileToEdit = "AppManifest.xml"
        $file = $zip.Entries.Where({$_.name -eq $fileToEdit})

        Write-Host "Read app manifest from '$file'."
        $appManifestFile = [System.IO.StreamReader]($file).Open()
        [xml]$xml = $appManifestFile.ReadToEnd()
        $appManifestFile.Close()

		 $appInformations = @{
                         ClientId = $xml.App.AppPrincipal.RemoteWebApplication.ClientId
                          Version = $xml.App.Version
               AllowAppOnlyPolicy = [bool]$xml.App.AppPermissionRequests.AllowAppOnlyPolicy
                        ProductID = $xml.App.ProductID
        }

		Write-Host "Read the following data from the app manifest:"
		Write-Host ($appInformations | Format-List | Out-String)
		

        return $appInformations

    }finally{
        # Write the changes and close the zip file
        $zip.Dispose()
    }
}

<#
.Synopsis
   Sets the app manifest data from the app package.  
.DESCRIPTION
   Sets the app manifest data from the app package. The Packages in unzipped and the app manifest xml is edited. 
   The following data can be set: client ID, app version, app web url
.EXAMPLE
   
#>
function Set-SpAddinPackageInformations {

    [CmdletBinding()]
    [OutputType([Hashtable])]
    param(		
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()]  
		[string]$Path, 

		[Parameter(Mandatory=$false, Position=1)]
		[string]$AppWebUrl,

		[Parameter(Mandatory=$false, Position=2)]
		[string]$AppVersion,

		[Parameter(Mandatory=$false, Position=3)]
		[string]$ClientId
	)

    # Open zip
    Add-Type -assembly  System.IO.Compression.FileSystem
    Write-Host "Open zip file '$Path'..."
    $zip =  [System.IO.Compression.ZipFile]::Open($Path, "Update")

    try{
        $fileToEdit = "AppManifest.xml"
        $file = $zip.Entries.Where({$_.name -eq $fileToEdit})

        Write-Host "Read app manifest from '$file'."
        $appManifestFile = [System.IO.StreamReader]($file).Open()
        [xml]$xml = $appManifestFile.ReadToEnd()
        $appManifestFile.Close()
		
		# change client id if available
        if (-not([string]::IsNullOrEmpty($ClientId))){
			Write-Host "Setting client id to $ClientId."
			$xml.App.AppPrincipal.RemoteWebApplication.ClientId = $ClientId  
        }
       
		#include app version if available
		if(-not([string]::IsNullOrEmpty($AppVersion))){
			Write-Host "Setting app version to $AppVersion."
			$xml.App.Version = $AppVersion
		}
       
		# Replace start URL
        if (-not([string]::IsNullOrEmpty($AppWebUrl))){            
            $value = $xml.App.Properties.StartPage
            Write-Host "Replace URL in '$value' with '$AppWebUrl'."
            $value = $value -replace "^.*\?","$($AppWebUrl)?" 
            $xml.App.Properties.StartPage = $value                       
        }

		# Save file
        Write-Host "Save manifest to '$file'."
        $appManifestFile = [System.IO.Stream]($file).Open()
        $appManifestFile.SetLength(0)
        $xml.Save($appManifestFile)
        $appManifestFile.Flush()
        $appManifestFile.Close()
        $appManifestFile.Dispose()

        $appManifestFile = [System.IO.StreamReader]($file).Open()
        [xml]$xml = $appManifestFile.ReadToEnd()
        Write-Host $xml
        $appManifestFile.Close()
        $appManifestFile.Dispose()

    }finally{
        # Write the changes and close the zip file
        $zip.Dispose()
    }
}

<#
.Synopsis
    Installs an app on SharePoint
.DESCRIPTION
    Installs an app on SharePoint. Sideloading feature will be activaed and deactivated before and after sideloading of the app. 
	The app ist installed on the connect web url.
.EXAMPLE
   Install-SpAddin -appPackage 'C:\path\to\file.app' -productId '045f93e8-5c9c-48a4-aaf1-396eff57dc04'
#>
function Install-SpAddin{
	
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()] 
		[System.IO.FileInfo]$AppPackage,

		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]  
		$ProductId
	)      
    $appName = [System.IO.Path]::GetFileNameWithoutExtension($AppPackage)
    $web =  Get-PnPWeb -Includes Language

    Write-Host "Start to install app $appName..."

    Write-Host "Installing app $appName..."
	$AppFilePath = [System.IO.Path]::GetFullPath($AppPackage)

	$appInstance = Import-PnPAppPackage -Path $AppFilePath -Locale $web.Language
   
    $appInstance = Wait-ForAppOperationComplete $appInstance.Id

    if (!$appInstance -Or $appInstance.Status -ne [Microsoft.SharePoint.Client.AppInstanceStatus]::Installed) 
    {
        if ($appInstance -And $appInstance.Id) 
        {
            Write-Error "App installation failed. To check app details, go to '$($web.Url.TrimEnd('/'))/_layouts/15/AppMonitoringDetails.aspx?AppInstanceId=$($appInstance.Id)'."
        }

        throw "App installation failed."
    }

    return $appInstance.Id
}

<#
.Synopsis
    
.DESCRIPTION
    
.EXAMPLE
   
#>
function Uninstall-SpAddin{
	param(
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()]  
		[string]$ProductId
	) 

	Write-Host "Searching for app with product id $($ProductId)..."
	$appInstances = Get-PnPAppInstance | Where-Object { $_.ProductId -eq $ProductId }

    if ($appInstances -And $appInstances.Length -gt 0) 
    {
        $appInstance = $appInstances[0]

        Write-Host "Uninstalling app with instance id $($appInstance.Id)..."
		Uninstall-PnPAppInstance -Identity $appInstance -Confirm:$false -Force -ErrorAction SilentlyContinue | Out-Null

        $appInstance = Wait-ForAppOperationComplete $appInstance.Id
        
        # Assume the app uninstallation succeeded
        Write-Host "App was uninstalled successfully."
    }
}

<#
.Synopsis
    
.DESCRIPTION
    
.EXAMPLE
   
#>
function Wait-ForAppOperationComplete{

	param(
		[Parameter(Mandatory=$true, Position=0)]
		[ValidateNotNullOrEmpty()] 
		[string]$AppInstanceId
	) 

    for ($i = 0; $i -le 200; $i++) 
    {
        try 
        {            
			$instance = Get-PnPAppInstance -Identity $AppInstanceId -ErrorAction SilentlyContinue         
        }
        catch
        {
            # When the uninstall finished, "app is not found" server exception will be thrown.
            # Assume the uninstalling operation succeeded.
            break
        }

        if (!$instance) 
        {
            break
        }

        $result = $instance.Status;
        if ($result -ne [Microsoft.SharePoint.Client.AppInstanceStatus]::Installed -And
            !$instance.InError -And 
            # If an app has failed to install correctly, it would return to initialized state if auto-cancel was enabled
            $result -ne [Microsoft.SharePoint.Client.AppInstanceStatus]::Initialized) 
        {
            Write-Host "Instance status: $result"
            Start-Sleep -m 1000
        }
        else 
        {
            break
        }
    }

    return $instance;
}