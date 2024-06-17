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

    Write-Host "PBIMonitor - Fetch Graph Started: $currentUTCtime"
    $config = Get-PBIMonitorConfig
    New-Item -ItemType Directory -Path ($config.OutputPath) -ErrorAction SilentlyContinue | Out-Null

    try {
        # Get the current universal time in the default string format.
        Write-Host "Starting Graph API Fetch"
        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()

        Add-Type -AssemblyName System.Web

        $rootOutputPath = "$($config.OutputPath)\graph"
        $outputPath = ("$rootOutputPath\{0:yyyy}\{0:MM}\{0:dd}" -f [datetime]::Today)

        # ensure folder
        New-Item -ItemType Directory -Path $outputPath -ErrorAction SilentlyContinue | Out-Null

        $graphUrl = "https://graph.microsoft.com/beta"
        $apiResource = "https://graph.microsoft.com"
        $graphCalls = @(
            @{
                GraphUrl = "$graphUrl/users?`$select=id,displayName,assignedLicenses,UserPrincipalName";
                FilePath = "$outputPath\users.json"
            },
            @{
                GraphUrl = "$graphUrl/subscribedSkus?`$select=id,capabilityStatus,consumedUnits,prepaidUnits,skuid,skupartnumber,prepaidUnits";
                FilePath = "$outputPath\subscribedskus.json"
            }
        )

        if ($config.GraphExtractGroups) {
            Write-Host "Adding graph call to extract groups"
            $graphCalls += @{
                #GraphUrl = "$graphUrl/groups?`$expand=members(`$select=id,displayName,appId,userPrincipalName)&`$select=id,displayName";
                GraphUrl = "$graphUrl/groups?`$filter=securityEnabled eq true&`$select=id,displayName";
                FilePath = "$outputPath\groups.json"
            }
        }

        $paginateCount = 10000

        if ($config.GraphPaginateCount) {
            $paginateCount = $config.GraphPaginateCount
        }

        Write-Host "GraphPaginateCount: $paginateCount"

        foreach ($graphCall in $graphCalls) {
            Write-Host "Getting OAuth token"
            $authToken = Get-AuthToken -resource $apiResource

            Write-Host "Calling Graph API: '$($graphCall.GraphUrl)'"
            $data = Read-FromGraphAPI -accessToken $authToken -url $graphCall.GraphUrl | Select-Object * -ExcludeProperty "@odata.id"

            $filePath = $graphCall.FilePath

            Get-ArrayInBatches -array $data -label "Read-FromGraphAPI Local Batch" -batchCount $paginateCount -script {
                param($dataBatch, $i)

                if ($i) {
                    $filePath = "$([System.IO.Path]::GetDirectoryName($filePath))\$([System.IO.Path]::GetFileNameWithoutExtension($filePath))_$i$([System.IO.Path]::GetExtension($filePath))"
                }

                if ($graphCall.GraphUrl -like "$graphUrl/groups*") {
                    Write-Host "Looping group batch to get members"
                    foreach($group in $dataBatch) {
                        $groupMembers = @(Read-FromGraphAPI -accessToken $authToken -url "$graphUrl/groups/$($group.id)/transitiveMembers?`$select=id,displayName,appId,userPrincipalName")
                        $group | Add-Member -NotePropertyName "members" -NotePropertyValue $groupMembers -ErrorAction SilentlyContinue
                    }
                }

                Write-Host "Writing to file: '$filePath'"
                ConvertTo-Json @($dataBatch) -Compress -Depth 5 | Out-File $filePath -Force

                Write-Host "Writing to Blob Storage"
                $storageRootPath = "$($config.StorageAccountContainerRootPath)/graph"
                $outputFilePath = $filePath

                if (Test-Path $outputFilePath) {
                    Add-FileToBlobStorage -storageAccountName $config.StorageAccountName -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $rootOutputPath
                    Remove-Item $outputFilePath -Force
                } else {
                    Write-Host "Cannot find file '$outputFilePath'"
                }
            }
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
