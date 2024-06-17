#Requires -Modules Az.Storage, MicrosoftPowerBIMgmt.Profile

param($Timer)

$global:erroractionpreference = 1
$reset = $false

$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
Import-Module "$currentPath\..\Utils.psm1" -Force

try {
    # Get the current universal time in the default string format.
    $currentUTCtime = (Get-Date).ToUniversalTime()

    if ($Timer.IsPastDue) {
        Write-Host "PowerShell timer is running late!"
    }

    Write-Host "PBIMonitor - Fetch Catalog Started: $currentUTCtime"

    $config = Get-PBIMonitorConfig

    New-Item -ItemType Directory -Path ($config.OutputPath) -ErrorAction SilentlyContinue | Out-Null
    $stateFilePath = "$($config.AppDataPath)\state.json"

    try {
        Write-Host "Starting Power BI Catalog Fetch"

        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()

        $outputPath = "$($config.OutputPath)\catalog"

        if (!$stateFilePath) {
            $stateFilePath = "$($config.OutputPath)\state.json"
        }

        if (Test-Path $stateFilePath) {
            $state = Get-Content $stateFilePath | ConvertFrom-Json

            #ensure mandatory fields
            $state | Add-Member -NotePropertyName "Catalog" -NotePropertyValue (new-object PSObject) -ErrorAction SilentlyContinue
            $state.Catalog | Add-Member -NotePropertyName "LastRun" -NotePropertyValue $null -ErrorAction SilentlyContinue
            $state.Catalog | Add-Member -NotePropertyName "LastFullScan" -NotePropertyValue $null -ErrorAction SilentlyContinue
        } else {
            $state = New-Object psobject
            $state | Add-Member -NotePropertyName "Catalog" -NotePropertyValue @{"LastRun" = $null; "LastFullScan" = $null } -Force
        }

        $state.Catalog.LastRun = [datetime]::UtcNow.Date.ToString("o")
        # ensure folders
        $scansOutputPath = Join-Path $outputPath ("scans\{0:yyyy}\{0:MM}\{0:dd}" -f [datetime]::Today)
        $snapshotOutputPath = Join-Path $outputPath ("snapshots\{0:yyyy}\{0:MM}\{0:dd}" -f [datetime]::Today)
        New-Item -ItemType Directory -Path $scansOutputPath -ErrorAction SilentlyContinue | Out-Null
        New-Item -ItemType Directory -Path $snapshotOutputPath -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Getting OAuth Token Managed Identity"
        $authToken = Get-AuthToken -resource 'https://api.fabric.microsoft.com/'

        #region ADMIN API

        $snapshotFiles = @()
        $filePath = "$snapshotOutputPath\apps.json"
        $snapshotFiles += $filePath

        if (!(Test-Path $filePath)) {
            Write-Host "Getting Power BI Apps List"
            $result = Invoke-WebRequest -Uri "https://api.powerbi.com/v1.0/myorg/admin/apps?`$top=5000&`$skip=0" -Headers @{'Authorization' = 'Bearer ' + $authToken } -RetryIntervalSec 5 | ConvertFrom-Json
            $result = @($result.value)

            if ($result.Count -ne 0) {
                ConvertTo-Json $result -Depth 10 -Compress | Out-File $filePath -force
            } else {
                Write-Host "Tenant without PowerBI apps"
            }
        } else {
            Write-Host "'$filePath' already exists"
        }

        # Save to Blob
        Write-Host "Writing Snapshots to Blob Storage"
        $storageRootPath = "$($config.StorageAccountContainerRootPath)/catalog"
        foreach ($outputFilePath in $snapshotFiles) {
            if (Test-Path $outputFilePath) {
                Add-FileToBlobStorage -storageAccountName $config.StorageAccountName -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $outputPath
                Remove-Item $outputFilePath -Force
            } else {
                Write-Warning "Cannot find file '$outputFilePath'"
            }
        }


        #endregion

        #region Workspace Scans: 1 - Get Modified; 2 - Start Scan for modified; 3 - Wait for scan finish; 4 - Get Results
        Write-Host "Getting workspaces to scan"
        $getInfoDetails = "lineage=true&datasourceDetails=true&getArtifactUsers=true&datasetSchema=false&datasetExpressions=false"
        if ($config.CatalogGetInfoParameters) {
            $getInfoDetails = $config.CatalogGetInfoParameters
        }

        $getModifiedWorkspacesParams = "excludePersonalWorkspaces=false&excludeInActiveWorkspaces=true"
        if ($config.CatalogGetModifiedParameters) {
            $getModifiedWorkspacesParams = $config.CatalogGetModifiedParameters
        }

        $fullScan = $false

        if ($state.Catalog.LastRun -and !$reset) {
            if (!($state.Catalog.LastRun -is [datetime])) {
                $state.Catalog.LastRun = [datetime]::Parse($state.Catalog.LastRun).ToUniversalTime()
            }

            if ($config.FullScanAfterDays) {
                if ($state.Catalog.LastFullScan) {
                    if (!($state.Catalog.LastFullScan -is [datetime])) {
                        $state.Catalog.LastFullScan = [datetime]::Parse($state.Catalog.LastFullScan).ToUniversalTime()
                    }

                    $daysSinceLastFullScan = ($state.Catalog.LastRun - $state.Catalog.LastFullScan).TotalDays
                    if ($daysSinceLastFullScan -ge $config.FullScanAfterDays) {
                        Write-Host "Triggering FullScan after $daysSinceLastFullScan days"
                        $fullScan = $true
                    } else {
                        Write-Host "Days to next fullscan: $($config.FullScanAfterDays - $daysSinceLastFullScan)"
                    }
                } else {
                    Write-Host "Triggering FullScan, because FullScanAfterDays is configured and LastFullScan is empty"
                    $fullScan = $true
                }
            }

            if (!$fullScan) {
                $modifiedSinceDate = $state.Catalog.LastRun

                $modifiedSinceDateMinDate = [datetime]::UtcNow.Date.AddDays(-30)

                if ($modifiedSinceDate -le $modifiedSinceDateMinDate) {
                    Write-Host "Last Run date was '$($modifiedSinceDate.ToString("o"))' but cannot go longer than 30 days"

                    $modifiedSinceDate = $modifiedSinceDateMinDate
                }

                $getModifiedWorkspacesParams = $getModifiedWorkspacesParams + "&modifiedSince=$($modifiedSinceDate.ToString("o"))"
            }

        } else {
            $fullScan = $true
        }

        $modifiedRequestUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/modified?$getModifiedWorkspacesParams"

        Write-Host "Reset: $reset"
        Write-Host "Since: $($state.Catalog.LastRun)"
        Write-Host "FullScan: $fullScan"
        Write-Host "Last FullScan: $($state.Catalog.LastFullScan)"
        Write-Host "FullScanAfterDays: $($config.FullScanAfterDays)"
        Write-Host "GetModified parameters '$getModifiedWorkspacesParams'"
        Write-Host "GetInfo parameters '$getInfoDetails'"

        # Get Modified Workspaces since last scan (Max 30 per hour)
        $workspacesModified = Invoke-WebRequest -Uri  $modifiedRequestUrl -Headers @{'Authorization' = 'Bearer ' + $authToken } -RetryIntervalSec 5 | ConvertFrom-Json

        if (!$workspacesModified -or $workspacesModified.Count -eq 0) {
            Write-Host "No workspaces modified"
        } else {
            Write-Host "Modified workspaces: $($workspacesModified.Count)"

            $throttleErrorSleepSeconds = 3700
            $scanStatusSleepSeconds = 5
            $getInfoOuterBatchCount = 1500
            $getInfoInnerBatchCount = 100

            Write-Host "Throttle Handling Variables: getInfoOuterBatchCount: $getInfoOuterBatchCount;  getInfoInnerBatchCount: $getInfoInnerBatchCount; throttleErrorSleepSeconds: $throttleErrorSleepSeconds"
            # postworkspaceinfo only allows 16 parallel requests, Get-ArrayInBatches allows to create a two level batch strategy. It should support initial load without throttling on tenants with ~50000 workspaces

            Get-ArrayInBatches -array $workspacesModified -label "GetInfo Global Batch" -batchCount $getInfoOuterBatchCount -script {
                param($workspacesModifiedOuterBatch, $i)
                $script:workspacesScanRequests = @()

                # Call GetInfo in batches of 100 (MAX 500 requests per hour)
                Get-ArrayInBatches -array $workspacesModifiedOuterBatch -label "GetInfo Local Batch" -batchCount $getInfoInnerBatchCount -script {
                    param($workspacesBatch, $x)
                    $bodyStr = @{"workspaces" = @($workspacesBatch.Id) } | ConvertTo-Json
                    $getInfoResult = Invoke-WebRequest -Uri  "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?$getInfoDetails" -Headers @{'Authorization' = 'Bearer ' + $authToken; "Content-Type" = 'application/json' } -Body $bodyStr -Method Post -RetryIntervalSec 5 | ConvertFrom-Json
                    $script:workspacesScanRequests += $getInfoResult
                }

                while (@($workspacesScanRequests | Where-Object status -in @("Running", "NotStarted"))) {
                    Write-Host "Waiting for scan results, sleeping for $scanStatusSleepSeconds seconds..."
                    Start-Sleep -Seconds $scanStatusSleepSeconds
                    foreach ($workspaceScanRequest in $workspacesScanRequests) {
                        $scanStatus = Invoke-WebRequest -Uri "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$($workspaceScanRequest.id)" -method Get -Headers @{'Authorization' = 'Bearer ' + $authToken; "Content-Type" = 'application/json' } -RetryIntervalSec 5 | ConvertFrom-Json
                        Write-Host "Scan '$($scanStatus.id)' : '$($scanStatus.status)'"
                        $workspaceScanRequest.status = $scanStatus.status
                    }
                }

                # Get Scan results (500 requests per hour) - https://docs.microsoft.com/en-us/rest/api/power-bi/admin/workspaceinfo_getscanresult
                foreach ($workspaceScanRequest in $workspacesScanRequests) {
                    $scanResult = Invoke-WebRequest -Uri  "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$($workspaceScanRequest.id)" -Headers @{'Authorization' = 'Bearer ' + $authToken; "Content-Type" = 'application/json' } -Method Get -RetryIntervalSec 5 | ConvertFrom-Json
                    Write-Host "Scan Result'$($scanStatus.id)' : '$($scanResult.workspaces.Count)'"
                    $fullScanSuffix = ""
                    if ($fullScan) {
                        $fullScanSuffix = ".fullscan"
                    }

                    $outputFilePath = "$scansOutputPath\$($workspaceScanRequest.id)$fullScanSuffix.json"
                    $scanResult | Add-Member –MemberType NoteProperty –Name "scanCreatedDateTime"  –Value $workspaceScanRequest.createdDateTime -Force
                    ConvertTo-Json $scanResult -Depth 10 -Compress | Out-File $outputFilePath -force

                    # Save to Blob
                    if (Test-Path $outputFilePath) {
                        Write-Host "Writing to Blob Storage"
                        $storageRootPath = "$($config.StorageAccountContainerRootPath)/catalog"
                        Add-FileToBlobStorage -storageAccountName $config.StorageAccountName -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $outputPath
                        Remove-Item $outputFilePath -Force
                    }
                }
            }

        }

        #endregion

        # Save State
        Write-Host "Saving state"
        New-Item -Path (Split-Path $stateFilePath -Parent) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        $state.Catalog.LastRun = [datetime]::UtcNow.Date.ToString("o")
        if ($fullScan) {
            $state.Catalog.LastFullScan = [datetime]::UtcNow.Date.ToString("o")
        }

        ConvertTo-Json $state | Out-File $stateFilePath -force -Encoding utf8
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
