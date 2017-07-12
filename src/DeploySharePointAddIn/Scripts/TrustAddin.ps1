[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
Param
(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]$AppInstanceId,

    [Parameter(Mandatory=$true, Position=1)]
    [ValidateNotNullOrEmpty()]
    [string]$WebUrl,

    [Parameter(Mandatory=$true, Position=2)]
    [ValidateNotNullOrEmpty()]
    [string]$DeployUserName,

    [Parameter(Mandatory=$true, Position=3)]
    [ValidateNotNullOrEmpty()]
    [string]$DeployPassword
)

<#
.Synopsis
    Waits as long as IE has not finished loading.
.DESCRIPTION
    Waits as long as IE has not finished loading. Checks in the IE COM-object if the Internet Explorer is still budy or not. 
	Time before initial check can be set. Default is one second.
.EXAMPLE
    WaitFor-IEReady -IE $ie -InitialWaitInSeconds 2
#>
function WaitFor-IEReady {

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [System.__ComObject]$IE,

        [Parameter(Mandatory=$false, Position=1)]
        [int]$InitialWaitInSeconds = 1
    )

    Start-Sleep -Seconds $InitialWaitInSeconds

    while ($IE.Busy) {
        Start-Sleep -milliseconds 100
    }
} 

<#
.Synopsis
    Invoke a JavaScript function.
.DESCRIPTION
    Use this function to run JavaScript on a web page. Your $Command can
    return a value which will be returned by this function unless $global
    switch is specified in which case $Command will be executed in global
    scope and cannot return a value. If you received error 80020101 it means
    you need to fix your JavaScript code.
.EXAMPLE
    Invoke-JavaScript -IE $ie -Command 'Post.IsSubmitReady();setTimeout(function() {Post.SubmitCreds(); }, 1000);'
.EXAMPLE
    $result = Invoke-JavaScript -IE $ie -Command 'Post.IsSubmitReady();setTimeout(function() {Post.SubmitCreds(); }, 1000);' -Global
#>
function Invoke-JavaScript {

    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [System.__ComObject]$IE,

        [Parameter(Mandatory=$true, Position=1)]
        [string]$Command,

		[Parameter(Mandatory=$false, Position=2)]
        [switch]$Global
    )

    if (-not $Global.IsPresent) {
        $Command = "document.body.setAttribute('PSResult', (function(){ $Command })());"
    }

    $document = $IE.document
    $window = $document.parentWindow
    $window.execScript($Command, 'javascript') | Out-Null

    if (-not $Global.IsPresent) {
        return $document.body.getAttribute('PSResult')
    }
}

<#
.Synopsis
    Trusts permission of an installed app.
.DESCRIPTION
    The function trusts an app after installation. This is done by loading the ie within powershel, 
	navigating to the page and pressing the permission button in the HTML-DOM with an invoked JavaScript function.
.EXAMPLE
    Trust-SPAddIn -AppInstanceId "d73cab34-b20a-46df-8457-aa0f22dc60da" -WebUrl https://my.sharepoint.company -UserName "user" -Password "password"
#>
function Trust-SPAddIn {

    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([int])]
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [guid]$AppInstanceId,

        [Parameter(Mandatory=$true, Position=1)]
        [string]$WebUrl,

        [parameter(Mandatory=$true, Position=2)] 
        [string]$UserName, 

        [parameter(Mandatory=$true, Position=3)] 
        [string]$Password
    )

    $authorizeURL = "$($WebUrl.TrimEnd('/'))/_layouts/15/appinv.aspx?AppInstanceId={$AppInstanceId}"

    [System.__ComObject]$ie = New-Object -com internetexplorer.application

    try
    {
        $ie.Visible = $false
        $ie.Navigate2($authorizeURL)

        WaitFor-IEReady $ie

        $docTitle = $ie.Document.Title
        if($docTitle -eq $null){
            $docTitle = $ie.LocationName
        }

        Write-Verbose $docTitle -Verbose

        if ($docTitle -match "Sign in to Office 365.*") {
        
            Write-Verbose "Authenticate $UserName to O365..."
            # Authorize against O365        
            $useAnotherLink = $ie.Document.getElementById("use_another_account_link")
            if ($useAnotherLink) {
            
                WaitFor-IEReady $ie
                $useAnotherLink.Click()
                WaitFor-IEReady $ie

            }

            $credUseridInputtext = $ie.Document.getElementById("cred_userid_inputtext")
            $credUseridInputtext.value = $UserName

            $credPasswordInputtext = $ie.Document.getElementById("cred_password_inputtext")
            $credPasswordInputtext.value = $Password
        
            WaitFor-IEReady $ie           

            # make a jQuery call
            $result = Invoke-JavaScript -IE $ie -Command "`nPost.IsSubmitReady();`nsetTimeout(function() {`nPost.SubmitCreds();`n}, 1000);"
     
            WaitFor-IEReady $ie -initialWaitInSeconds 5            
        }

        $docTitle = $ie.Document.Title
        if($docTitle -eq $null){
            $docTitle = $ie.LocationName
        }

        Write-Verbose $docTitle -Verbose
        
        if ($docTitle -match "Do you trust.*")
        {
            Start-Sleep -seconds 5

            $button = $ie.Document.getElementById("ctl00_PlaceHolderMain_BtnAllow")

			if ($button -eq $null) {
				$button = $ie.Document.getElementById("ctl00_PlaceHolderMain_LnkRetrust")
			}
            
            if ($button -eq $null) {

                throw "Could not find button to press"

            }else{

                $button.click()
 
                WaitFor-IEReady $ie

                $docTitle = $ie.Document.Title
                if($docTitle -eq $null){
                    $docTitle = $ie.LocationName
                }

                #if the button press was successful, we should now be on the Site Settings page.. 
                if ($docTitle -like "*trust*") {

                    throw "Error: $($ie.Document.body.getElementsByClassName("ms-error").item().InnerText)"

                }else{

                    Write-Verbose "App was trusted successfully!"
                }
            }

        }else{

            throw "Unexpected page '$($ie.LocationName)' was loaded. Please check your url."
        }
    }
    finally
    {
        $ie.Quit()
    } 
}  

Trust-SPAddIn -AppInstanceId $AppInstanceId -WebUrl $WebUrl -UserName $DeployUserName -Password $DeployPassword
