
$extensionFileJson = Get-Content -Path '.\vss-extension.json' | Out-String | ConvertFrom-Json

#copy only to used extension paths
$extensionIds = $extensionFileJson.contributions.id

$extensionIds | ForEach-Object {

    $taskIdName = $_

    $destinationFolder = ".\$taskIdName\ps_modules"

    #remove any content from those folder, as they are temporary
    Remove-Item -Path $destinationFolder -Recurse -Force
    #copy and overwrite all
    Copy-Item -Path ".\ps_modules" -Destination $destinationFolder -Recurse -Force


}