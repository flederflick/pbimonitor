$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
Import-Module "$currentPath\AzureFunction\Utils.psm1" -Force

$authToken = Get-AuthToken -resource 'https://api.fabric.microsoft.com/'


$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"

try {
    $stopwatch = [System.Diagnostics.Stopwatch]::new()
    $stopwatch.Start()

    $currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)

    Set-Location $currentPath

    if (Test-Path $configFilePath) {
        $config = Get-Content $configFilePath | ConvertFrom-Json
    } else {
        throw "Cannot find config file '$configFilePath'"
    }

    Write-Host "Getting OAuth Token for ServicePrincipal to find the ObjectId"

    $result = Invoke-WebRequest -Uri $url -Headers @{'Authorization' = 'Bearer ' + $tokenFMI } | ConvertFrom-Json

    Write-Host "Apps Returned: $($result.value.Count)"

} finally {
    $stopwatch.Stop()

    Write-Host "Elapsed: $($stopwatch.Elapsed.TotalSeconds)s"
}
