param(
  [Parameter(Mandatory)]
  [ValidateNotNullOrEmpty()]
  [string]$AgentToolPath = "C:\temp"
)

$here = (Split-Path -Parent $MyInvocation.MyCommand.Path)
. $here\..\DeployArtifactsWithPnP\Scripts\Utility.ps1



if($AgentToolPath -eq $null){
 $AgentToolPath = "$($env:AGENT_RELEASEDIRECTORY)\_temp"
}

Describe 'Utility Tests' {

    Context -Name "Call PnP libraries by environment"{

        It -Name "Given valid -Name <Environment>, it returns '<Expected>'"  -TestCases @(
            @{Environment = "SpOnline"; Expected = $true}
            @{Environment = "Sp2016"; Expected = $true}
            @{Environment = "Sp2013"; Expected = $true}
        ){
            param ($Environment, $Expected)            

            $isLoaded = Load-PnPPackages -SharePointVersion $Environment -AgentToolPath $AgentToolPath

            Write-Host $Expected

            $isLoaded | Should -Be $Expected
        }

        It -Name "Given invalid parameter -Name 'SharePointOffline', it return `$false" {

            { Load-PnPPackages -SharePointVersion "SharePointOffline" -AgentToolPath $AgentToolPath } | Should -Throw 'Only SharePoint 2013, 2016 or SharePoint Online are supported at the moment'

        }

    }   

}