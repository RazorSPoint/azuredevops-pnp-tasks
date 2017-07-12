[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
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
    [string]$DeployUserName,

    [Parameter(Mandatory=$true, Position=3)]
    [ValidateNotNullOrEmpty()]
    [string]$DeployPassword,
		
    [Parameter(Mandatory=$true, Position=4)]
    [ValidateSet('SpOnline','Sp2016','Sp2013')]
    [string]$SharePointVersion,

	[Parameter(Mandatory=$false, Position=5)]
    [ValidateNotNullOrEmpty()]
    [string]$Machine = $null,

	[Parameter(Mandatory=$false, Position=6)]
    [string]$AppVersion,

	[Parameter(Mandatory=$false, Position=7)]
    [switch]$ReinstallIfExisting

)

Load-Assemblies $SharePointVersion

$secpasswd = ConvertTo-SecureString $DeployPassword -AsPlainText -Force
$adminCredentials = New-Object System.Management.Automation.PSCredential ($DeployUserName, $secpasswd)

Connect-PnPOnline -Url $WebUrl -Credentials $adminCredentials
Write-Host "Successfully connected to '$WebUrl'..."

Set-SpAddinPackageInformations -Path $AppFilePath -AppVersion $AppVersion

Write-Host "Get package informations"
$appPackageInformations = Get-SpAddinPackageInformations -Path $AppFilePath

if($ReinstallIfExisting){
	Uninstall-SpAddin $appPackageInformations.ProductId
}

Write-Host "Update SharePoint app in '$WebUrl'"

#enable sideloading feature
Enable-SideLoading -Enable $true -SharePointVersion $SharePointVersion -Machine $Machine -DeployUserName $DeployUserName -DeployPassword $DeployPassword -Force

#sideload app
Write-Host "installing app from path $AppFilePath"
$appInformation = Import-PnPAppPackage -Path $AppFilePath

#deactivate sideloading feature
Enable-SideLoading -Enable $false -SharePointVersion $SharePointVersion -Machine $Machine -DeployUserName $DeployUserName -DeployPassword $DeployPassword -Force

Write-Host "Done."