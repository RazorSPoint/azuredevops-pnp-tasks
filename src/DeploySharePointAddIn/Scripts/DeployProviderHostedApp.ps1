[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
[OutputType([int])]
Param
(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$AppFilePath,

    [Parameter(Mandatory=$true, Position=1)]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrl,

    [Parameter(Mandatory=$true, Position=2)]
    [ValidateNotNullOrEmpty()]
    [string]$RemoteAppUrl,

    [Parameter(Mandatory=$true, Position=3)]
    [ValidateNotNullOrEmpty()]
    [string]$DeployUserName,

    [Parameter(Mandatory=$true, Position=4)]
    [ValidateNotNullOrEmpty()]
    [string]$DeployPassword,

    [Parameter(Mandatory=$true, Position=5)]
    [ValidateSet('SpOnline','Sp2016','Sp2013')]
    [string]$SharePointVersion,

	[Parameter(Mandatory=$false, Position=6)]
    [string]$Machine = $null,

	[Parameter(Mandatory=$false, Position=7)]
    [string]$AppVersion,

	[Parameter(Mandatory=$false, Position=8)]
    [string]$ClientId,

	[Parameter(Mandatory=$false, Position=9)]
    [switch]$ReinstallIfExisting
)

Load-Assemblies $SharePointVersion

$secpasswd = ConvertTo-SecureString $DeployPassword -AsPlainText -Force
$adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, $secpasswd)

Write-Host "Connect to '$WebUrl' as '$DeployUserName'..."
Connect-PnPOnline -Url $WebUrl -Credentials $adminCredentials
Write-Host "Successfully connected to '$WebUrl'..."    

Write-Host "Get package informations"
$appPackageInformations = Get-SpAddinPackageInformations -Path $AppFilePath

if($ReinstallIfExisting){	
    Uninstall-SpAddin $appPackageInformations.ProductId
}

Set-SpAddinPackageInformations -Path $AppFilePath -AppWebUrl $RemoteAppUrl -AppVersion $AppVersion -ClientId $ClientId

Enable-SideLoading -Enable $true -SharePointVersion $SharePointVersion -Machine $Machine -DeployUserName $DeployUserName -DeployPassword $DeployPassword -Force

Write-Host "Update SharePoint app in '$WebUrl' to Version '$($appPackageInformations.Version)'..."
$appId = Install-SpAddin -appPackage $AppFilePath -productId $appPackageInformations.ProductID
Write-Host "Done."

Enable-SideLoading -Enable $false -SharePointVersion $SharePointVersion -Machine $Machine -DeployUserName $DeployUserName -DeployPassword $DeployPassword -Force
