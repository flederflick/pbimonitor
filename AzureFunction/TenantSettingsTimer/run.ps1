#Requires -Modules Az.Storage

param($Timer)

$global:erroractionpreference = 1
$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
Import-Module "$currentPath\..\Utils.psm1" -Force
try {

    # Get the current universal time in the default string format.
    $currentUTCtime = (Get-Date).ToUniversalTime()

    if ($Timer.IsPastDue) {
        Write-Host "PowerShell timer is running late!"
    }

    Write-Host "PBIMonitor - Fetch Tenant Settings Started: $currentUTCtime"
    $config = Get-PBIMonitorConfig
    New-Item -ItemType Directory -Path ($config.OutputPath) -ErrorAction SilentlyContinue | Out-Null

    try {

        Write-Host "Starting Tenant API Fetch"
        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()

        Add-Type -AssemblyName System.Web

        $rootOutputPath = "$($config.OutputPath)\tenant"
        $outputPath = ("$rootOutputPath\{0:yyyy}\{0:MM}\{0:dd}" -f [datetime]::Today)

        New-Item -ItemType Directory -Path $outputPath -ErrorAction SilentlyContinue | Out-Null

        $tenantUrl = "https://api.fabric.microsoft.com/v1/admin/tenantsettings"
        $apiResource = "https://api.fabric.microsoft.com/"
        $TenantFilePath = "$($outputPath)\tenant-settings.json"

        Write-Host "Getting OAuth token"
        $authToken = Get-AuthToken -resource $apiResource

        Write-Host "Calling Graph API: https://api.fabric.microsoft.com/v1/admin/tenantsettings"
        $data = Read-FromTenantAPI -accessToken $authToken -url $tenantUrl

        Write-Host "Writing to file: '$($TenantFilePath)'"
        ConvertTo-Json $data -Compress -Depth 5 | Out-File $TenantFilePath -Force

        Write-Host "Writing to Blob Storage"
        $storageRootPath = "$($config.StorageAccountContainerRootPath)/tenant"

        $outputFilePath = $TenantFilePath
        if (Test-Path $outputFilePath) {
            Add-FileToBlobStorage -storageAccountName $config.StorageAccountName -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $rootOutputPath
            Remove-Item $outputFilePath -Force
        } else {
            Write-Host "Cannot find file '$outputFilePath'"
        }
    } finally {
        $stopwatch.Stop()
        Write-Host "Elapsed: $($stopwatch.Elapsed.TotalSeconds)s"
    }

    Write-Host "End"
} catch {

    $ex = $_.Exception
    if ($ex.ToString().Contains("429 (Too Many Requests)")) {
        throw "429 Throthling Error - Need to wait before making another request..."
    }
    Resolve-PowerBIError -Last
    throw
}
