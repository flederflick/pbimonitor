#Requires -Modules Az.Storage,MicrosoftPowerBIMgmt.Profile, MicrosoftPowerBIMgmt.Workspaces

param($Timer)

$global:erroractionpreference = 1
$workspaceFilter = @()
$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
Import-Module "$currentPath\..\Utils.psm1" -Force
try {
    # Get the current universal time in the default string format.
    $currentUTCtime = (Get-Date).ToUniversalTime()

    if ($Timer.IsPastDue) {
        Write-Host "PowerShell timer is running late!"
    }

    Write-Host "PBIMonitor - Dataset Refresh Started: $currentUTCtime"

    $config = Get-PBIMonitorConfig
    New-Item -ItemType Directory -Path $config.OutputPath -ErrorAction SilentlyContinue | Out-Null

    try {
        Write-Host "Starting Power BI Dataset Refresh History Fetch"

        $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()

        # ensure folder
        $rootOutputPath = "$($config.OutputPath)\datasetrefresh"
        $outputPath = ("$rootOutputPath\{0:yyyy}\{0:MM}\{0:dd}" -f [datetime]::Today)
        $tempPath = Join-Path $outputPath "_temp"

        New-Item -ItemType Directory -Path $tempPath -ErrorAction SilentlyContinue | Out-Null
        New-Item -ItemType Directory -Path $outputPath -ErrorAction SilentlyContinue | Out-Null

        Write-Host "Getting OAuth Token Managed Identity"
        $authToken = Get-AuthToken -resource 'https://api.fabric.microsoft.com/'

        # Find Token Object Id, by decoding OAUTH TOken - https://blog.kloud.com.au/2019/07/31/jwtdetails-powershell-module-for-decoding-jwt-access-tokens-with-readable-token-expiry-time/
        $tokenPayload = $authToken.Split(".")[1].Replace('-', '+').Replace('_', '/')
        while ($tokenPayload.Length % 4) { $tokenPayload += "=" }
        $tokenPayload = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tokenPayload)) | ConvertFrom-Json
        $pbiUserIdentifier = $tokenPayload.oid

        #region Workspace Users

        # Get workspaces + users
        $workspacesFilePath = "$tempPath\workspaces.datasets.json"

        if (!(Test-Path $workspacesFilePath)) {
            # $workspaces = Get-PowerBIWorkspace -Scope Organization -All -Include Datasets
            $count = 0
            $workspaces = New-Object System.Collections.ArrayList
            do {
                $count = $count + 5000
                $activityAPIUrl = "https://api.powerbi.com/v1.0/myorg/admin/groups?`$top=5000&`$skip=$($count - 5000)"
                $result = Invoke-WebRequest -Uri $activityAPIUrl -Headers @{'Authorization' = 'Bearer ' + $authToken } -SkipHttpErrorCheck -RetryIntervalSec 5 | ConvertFrom-Json


                foreach($workspace in $result.value) {
                    if($workspace.state -eq "Active" -and $workspace.type -eq "Workspace") {
                        $url = "https://api.powerbi.com/v1.0/myorg/admin/groups/$($workspace.id)/datasets"
                        $resultWorkspace = Invoke-WebRequest -Uri $url -Headers @{'Authorization' = 'Bearer ' + $authToken } -SkipHttpErrorCheck -RetryIntervalSec 5 | ConvertFrom-Json
                        $workspace | Add-Member -Name Datasets -MemberType NoteProperty -Value $resultWorkspace.value
                    }
                    $workspaces.Add($workspace)
                }
            }until($count -ge $result.'@odata.count')
            $workspaces | ConvertTo-Json -Depth 5 -Compress | Out-File $workspacesFilePath
        } else {
            Write-Host "Workspaces file already exists"
            $workspaces = Get-Content -Path $workspacesFilePath | ConvertFrom-Json
        }
        Write-Host "Workspaces: $($workspaces.Count)"

        $workspaces = $workspaces | Where-Object { $_.users | Where-Object { $_.identifier -ieq $pbiUserIdentifier } }
        Write-Host "Workspaces where user is a member: $($workspaces.Count)"

        # Only look at Active, V2 Workspaces and with Datasets
        $workspaces = @($workspaces | Where-Object { $_.type -eq "Workspace" -and $_.state -eq "Active" -and $_.datasets.Count -gt 0 })

        if ($workspaceFilter -and $workspaceFilter.Count -gt 0) {
            $workspaces = @($workspaces | Where-Object { $workspaceFilter -contains $_.Id })
        }

        Write-Host "Workspaces to get refresh history: $($workspaces.Count)"

        $total = $Workspaces.Count
        $item = 0

        foreach ($workspace in $Workspaces) {
            $item++
            Write-Host "Processing workspace: '$($workspace.Name)' $item/$total"
            Write-Host "Datasets: $(@($workspace.datasets).Count)"
            $refreshableDatasets = @($workspace.datasets | Where-Object { $_.isRefreshable -eq $true -and $_.addRowsAPIEnabled -eq $false })
            Write-Host "Refreshable Datasets: $($refreshableDatasets.Count)"

            foreach ($dataset in $refreshableDatasets) {
                try {
                    Write-Host "Processing dataset: '$($dataset.name)'"
                    Write-Host "Getting refresh history"
                    # $dsRefreshHistory = Invoke-PowerBIRestMethod -Url "groups/$($workspace.id)/datasets/$($dataset.id)/refreshes" -Method Get | ConvertFrom-Json
                    $dsRefreshHistory = Invoke-WebRequest -Uri "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.id)/datasets/$($dataset.id)/refreshes" -Headers @{'Authorization' = 'Bearer ' + $authToken } -RetryIntervalSec 5 | ConvertFrom-Json
                    $dsRefreshHistory = $dsRefreshHistory.value

                    if ($dsRefreshHistory) {
                        $dsRefreshHistory = @($dsRefreshHistory | Select-Object *, @{Name = "dataSetId"; Expression = { $dataset.id } }, @{Name = "dataSet"; Expression = { $dataset.name } }`
                                , @{Name = "group"; Expression = { $workspace.name } }, @{Name = "configuredBy"; Expression = { $dataset.configuredBy } })

                        $dsRefreshHistoryGlobal += $dsRefreshHistory
                    }
                } catch {
                    $ex = $_.Exception
                    Write-Error -message "Error processing dataset: '$($ex.Message)'" -ErrorAction Continue
                    # If its unauthorized no need to advance to other datasets in this workspace
                    if ($ex.Message.Contains("Unauthorized") -or $ex.Message.Contains("(404) Not Found")) {
                        Write-Host "Got unauthorized/notfound, skipping workspace"

                        break

                    }
                }
            }
        }

        if ($dsRefreshHistoryGlobal.Count -gt 0) {
            $outputFilePath = "$outputPath\workspaces.datasets.refreshes.json"
            ConvertTo-Json @($dsRefreshHistoryGlobal) -Compress -Depth 5 | Out-File $outputFilePath -force

            if (Test-Path $outputFilePath) {
                Write-Host "Writing to Blob Storage"
                $storageRootPath = "$($config.StorageAccountContainerRootPath)/datasetrefresh"
                Add-FileToBlobStorage -storageAccountName $config.StorageAccountName -storageContainerName $config.StorageAccountContainerName -storageRootPath $storageRootPath -filePath $outputFilePath -rootFolderPath $rootOutputPath
                Remove-Item $outputFilePath -Force
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
