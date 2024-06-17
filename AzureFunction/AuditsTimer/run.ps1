#Requires -Modules Az.Storage,MicrosoftPowerBIMgmt.Profile

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

    Write-Host "PBIMonitor - Fetch Activity Started: $currentUTCtime"
    $config = Get-PBIMonitorConfig

    New-Item -ItemType Directory -Path ($config.OutputPath) -ErrorAction SilentlyContinue | Out-Null
    $stateFilePath = "$($config.AppDataPath)\state.json"

    try {
        Write-Host "Starting Power BI Activity Fetch"

        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()

        if ($config.ActivityFileBatchSize) {
            $outputBatchCount = $config.ActivityFileBatchSize
        } else {
            $outputBatchCount = 5000
        }

        $rootOutputPath = "$($config.OutputPath)\activity"
        New-Item -ItemType Directory -Path $rootOutputPath -ErrorAction SilentlyContinue | Out-Null

        $outputPath = "$rootOutputPath\{0:yyyy}\{0:MM}"

        if (!$stateFilePath) {
            $stateFilePath = "$($config.OutputPath)\state.json"
        }

        if (Test-Path $stateFilePath) {
            $state = Get-Content $stateFilePath | ConvertFrom-Json
        } else {
            $state = New-Object psobject
        }

        $maxHistoryDate = [datetime]::UtcNow.Date.AddDays(-30)

        if ($state.Activity.LastRun) {
            if (!($state.Activity.LastRun -is [datetime])) {
                $state.Activity.LastRun = [datetime]::Parse($state.Activity.LastRun).ToUniversalTime()
            }
            $pivotDate = $state.Activity.LastRun
        } else {
            $state | Add-Member -NotePropertyName "Activity" -NotePropertyValue @{"LastRun" = $null } -Force
            $pivotDate = $maxHistoryDate
        }

        if ($pivotDate -lt $maxHistoryDate) {
            Write-Host "Last run was more than 30 days ago"
            $pivotDate = $maxHistoryDate
        }

        Write-Host "Since: $($pivotDate.ToString("s"))"
        Write-Host "OutputBatchCount: $outputBatchCount"

        Write-Host "Getting OAuth Token Managed Identity"
        $authToken = Get-AuthToken -resource 'https://api.fabric.microsoft.com/'

        # Gets audit data for each day
        while ($pivotDate -le [datetime]::UtcNow) {
            Write-Host "Getting audit data for: '$($pivotDate.ToString("yyyyMMdd"))'"

            $activityAPIUrl = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$($pivotDate.ToString("s"))'&endDateTime='$($pivotDate.AddHours(24).AddSeconds(-1).ToString("s"))'"

            $audits = @()
            $pageIndex = 1
            $flagNoActivity = $true

            do {
                if (!$result.continuationUri) {
                    # $result = Invoke-PowerBIRestMethod -Url $activityAPIUrl -method Get | ConvertFrom-Json
                    $result = Invoke-WebRequest -Uri $activityAPIUrl -Headers @{'Authorization' = 'Bearer ' + $authToken } -RetryIntervalSec 5 | ConvertFrom-Json
                } else {
                    # $result = Invoke-PowerBIRestMethod -Url $result.continuationUri -method Get | ConvertFrom-Json
                    $result = Invoke-WebRequest -Uri $result.continuationUri -Headers @{'Authorization' = 'Bearer ' + $authToken } -RetryIntervalSec 5 | ConvertFrom-Json
                }

                if ($result.activityEventEntities) {
                    $audits += @($result.activityEventEntities)
                }

                if ($audits.Count -ne 0 -and ($audits.Count -ge $outputBatchCount -or $null -eq $result.continuationToken)) {
                    # To avoid duplicate data on existing files, first dont append pageindex to overwrite existing full file

                    if ($pageIndex -eq 1) {
                        $outputFilePath = ("$outputPath\{0:yyyyMMdd}.json" -f $pivotDate)
                    } else {
                        $outputFilePath = ("$outputPath\{0:yyyyMMdd}_$pageIndex.json" -f $pivotDate)
                    }

                    Write-Host "Writing '$($audits.Count)' audits"
                    New-Item -Path (Split-Path $outputFilePath -Parent) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                    ConvertTo-Json @($audits) -Compress -Depth 10 | Out-File $outputFilePath -force

                    if (Test-Path $outputFilePath) {
                        Write-Host "Writing to Blob Storage"
                        $storageRootPath = "$($config.StorageAccountContainerRootPath)/activity"
                        Add-FileToBlobStorage -storageAccountName $config.StorageAccountName  -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $rootOutputPath

                        Write-Host "Deleting local file '$outputFilePath'"
                        Remove-Item $outputFilePath -Force
                    }

                    $flagNoActivity = $false
                    $pageIndex++
                    $audits = @()
                }
            }
            while($null -ne $result.continuationToken)

            if ($flagNoActivity) {
                Write-Warning "No audit logs for date: '$($pivotDate.ToString("yyyyMMdd"))'"
            }

            $state.Activity.LastRun = $pivotDate.Date.ToString("o")
            $pivotDate = $pivotDate.AddDays(1)

            # Save state
            Write-Host "Saving state"
            New-Item -Path (Split-Path $stateFilePath -Parent) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
            ConvertTo-Json $state | Out-File $stateFilePath -force -Encoding utf8
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
