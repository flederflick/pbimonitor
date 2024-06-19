function Get-PBIMonitorConfig {
    Write-Host "Building PBIMonitor Config from Azure Function Configuration"

    $appDataPath = $env:PBIMONITOR_AppDataPath
    if (!$appDataPath) {
        $appDataPath = "C:\home\data\pbimonitor"
    }

    $outputPath = $env:PBIMONITOR_DataPath
    if (!$outputPath) {
        $outputPath = "$($env:temp)\PBIMonitorData\$([guid]::NewGuid().ToString("n"))"
    }

    $scriptsPath = $env:PBIMONITOR_ScriptsPath
    if (!$scriptsPath) {
        $scriptsPath = "C:\home\site\wwwroot\Scripts"
    }

    $environment = $env:PBIMONITOR_ServicePrincipalEnvironment
    if (!$environment) {
        $environment = "Public"
    }

    $containerName = $env:PBIMONITOR_StorageContainerName
    if (!$containerName) {
        $containerName = "pbimonitor"
    }

    $storageaccountname = $env:PBIMONITOR_StorageAccountName;
    if (!$storageaccountname) {
        $storageaccountname = $env:AzureWebJobsStorage__accountname;
    }

    $containerRootPath = $env:PBIMONITOR_StorageRootPath
    if (!$containerRootPath) {
        $containerRootPath = "raw"
    }

    $flagExtractGroups = $false

    if($env:PBIMONITOR_GraphExtractGroups) {
        $flagExtractGroups = [System.Convert]::ToBoolean($env:PBIMONITOR_GraphExtractGroups)
    }

    $config = @{
        "AppDataPath"                     = $appDataPath;
        "ScriptsPath"                     = $scriptsPath;
        "OutputPath"                      = $outputPath;
        "StorageAccountName"              = $storageaccountname;
        "StorageAccountContainerName"     = $containerName;
        "StorageAccountContainerRootPath" = $containerRootPath;
        "ActivityFileBatchSize"           = $env:PBIMONITOR_ActivityFileBatchSize;
        "FullScanAfterDays"               = $env:PBIMONITOR_FullScanAfterDays;
        "CatalogGetInfoParameters"        = $env:PBIMONITOR_CatalogGetInfoParameters;
        "CatalogGetModifiedParameters"    = $env:PBIMONITOR_CatalogGetModifiedParameters;
        "GraphPaginateCount"              = $env:PBIMONITOR_GraphPaginateCount;
        "GraphExtractGroups"              = $flagExtractGroups;
        "ServicePrincipal"                = @{
            "TenantId"    = $env:PBIMONITOR_ServicePrincipalTenantId;
            "Environment" = $environment;
        }
    }

    Write-Host "AppDataPath: $appDataPath"
    Write-Host "ScriptsPath: $scriptsPath"
    Write-Host "OutputPath: $outputPath"

    Write-Output $config

}

function Add-FolderToBlobStorage {
    [cmdletbinding()]
    param
    (
        [string]$storageAccountName,
        [string]$storageContainerName,
        [string]$storageRootPath,
        [string]$folderPath,
        [string]$rootFolderPath,
        [bool]$ensureContainer = $true
    )

    $ctx = New-AzStorageContext -StorageAccountName $storageAccountName

    if ($ensureContainer) {
        Write-Host "Ensuring container '$storageContainerName'"

        New-AzStorageContainer -Context $ctx -Name $storageContainerName -Permission Off -ErrorAction SilentlyContinue | Out-Null
    }

    $files = @(Get-ChildItem -Path $folderPath -Filter *.* -Recurse -File)

    Write-Host "Adding folder '$folderPath' (files: $($files.Count)) to blobstorage '$storageAccountName/$storageContainerName/$storageRootPath'"

    if (!$rootFolderPath) {
        $rootFolderPath = $folderPath
    }

    foreach ($file in $files) {
        $filePath = $file.FullName

        Add-FileToBlobStorageInternal -ctx $ctx -filePath $filePath -storageRootPath $storageRootPath -rootFolderPath  $rootFolderPath
    }
}

function Add-FileToBlobStorage {
    [cmdletbinding()]
    param
    (
        [string]$storageAccountName,
        [string]$storageContainerName,
        [string]$storageRootPath,
        [string]$filePath,
        [string]$rootFolderPath,
        [bool]$ensureContainer = $true
    )

    $ctx = New-AzStorageContext -StorageAccountName $storageAccountName

    if ($ensureContainer) {
        New-AzStorageContainer -Context $ctx -Name $storageContainerName -Permission Off -ErrorAction SilentlyContinue | Out-Null
    }

    Add-FileToBlobStorageInternal -ctx $ctx -filePath $filePath -storageRootPath $storageRootPath -rootFolderPath $rootFolderPath

}

function Add-FileToBlobStorageInternal {
    param
    (
        $ctx,
        [string]$storageRootPath,
        [string]$filePath,
        [string]$rootFolderPath
    )

    if (Test-Path $filePath) {
        Write-Host "Adding file '$filePath' files to blobstorage '$storageAccountName/$storageContainerName/$storageRootPath'"
        $filePath = Resolve-Path $filePath
        $filePath = $filePath.ToLower()

        if ($rootFolderPath) {
            $rootFolderPath = Resolve-Path $rootFolderPath
            $rootFolderPath = $rootFolderPath.ToLower()

            $fileName = (Split-Path $filePath -Leaf)
            $parentFolder = (Split-Path $filePath -Parent)
            $relativeFolder = $parentFolder.Replace($rootFolderPath, "").Replace("\", "/").TrimStart("/").Trim();
        }

        if (!([string]::IsNullOrEmpty($relativeFolder))) {
            $blobName = "$storageRootPath/$relativeFolder/$fileName"
        } else {
            $blobName = "$storageRootPath/$fileName"
        }

        Set-AzStorageBlobContent -File $filePath -Container $storageContainerName -Blob $blobName -Context $ctx -Force | Out-Null
    } else {
        Write-Host "File '$filePath' dont exist"
    }
}

function Get-ArrayInBatches {
    [cmdletbinding()]
    param
    (
        [array]$array,
        [int]$batchCount,
        [ScriptBlock]$script,
        [string]$label = "Get-ArrayInBatches"
    )

    $skip = 0
    $i = 0

    do {
        $batchItems = @($array | Select-Object -First $batchCount -Skip $skip)

        if ($batchItems) {
            Write-Host "[$label] Batch: $($skip + $batchCount) / $($array.Count)"
            Invoke-Command -ScriptBlock $script -ArgumentList @($batchItems, $i)
            $skip += $batchCount
        }
        $i++
    }
    while($batchItems.Count -ne 0 -and $batchItems.Count -ge $batchCount)
}

function Wait-On429Error {
    [cmdletbinding()]
    param
    (
        [ScriptBlock]$script,
        [int]$sleepSeconds = 3601,
        [int]$tentatives = 1
    )

    try {
        Invoke-Command -ScriptBlock $script
    } catch {

        $ex = $_.Exception

        $errorText = $ex.ToString()
        ## If code errors at this location it is likely due to a 429 error. The PowerShell comandlets do not handle 429 errors with the appropriate message. This code will cover the known errors codes.
        if ($errorText -like "*Error reading JObject from JsonReader*" -or ($errorText -like "*429 (Too Many Requests)*" -or $errorText -like "*Response status code does not indicate success: *" -or $errorText -like "*You have exceeded the amount of requests allowed*")) {

            Write-Host "'429 (Too Many Requests)' Error - Sleeping for $sleepSeconds seconds before trying again" -ForegroundColor Yellow
            Write-Host "Printing Error for Logs: '$($errorText)'"
            $tentatives = $tentatives - 1

            if ($tentatives -lt 0) {
                throw "[Wait-On429Error] Max Tentatives reached!"
            } else {
                Start-Sleep -Seconds $sleepSeconds

                Wait-On429Error -script $script -sleepSeconds $sleepSeconds -tentatives $tentatives
            }
        } else {
            throw
        }
    }
}

function Get-AuthToken() {
    [cmdletbinding()]
    param
    (
        [string]$resource
    )

    $accesstoken = Get-AuthTokenMI -resource $resource
    # $accesstoken = Get-AuthTokenSPN -resource $resource
    write-output $accesstoken
}

function Get-AuthTokenMI() {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)][Validateset('https://graph.microsoft.com/', 'https://graph.microsoft.com', 'https://analysis.windows.net/powerbi/api', 'https://analysis.windows.net/powerbi/api/', 'https://api.fabric.microsoft.com/', 'https://api.fabric.microsoft.com')][string]$resource
    )

    write-host "Getting token for resource: $resource"
    $headers = @{"X-IDENTITY-HEADER" = $env:IDENTITY_HEADER }
    $ProgressPreference = "SilentlyContinue"
    $response = Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resource&api-version=2019-08-01" -Headers $headers -RetryIntervalSec 5
    $token = ConvertFrom-Json $response
    $accesstoken = $token.access_token
    write-output $accesstoken
}

function Get-AuthTokenSPN {
    [cmdletbinding()]
    param
    (
        [string]$authority = "https://login.microsoftonline.com",
        [string]$tenantid = $env:PBIMONITOR_ServicePrincipalTenantId,
        [string]$appid = $env:PBIMONITOR_ServicePrincipalId,
        [string]$appsecret = $env:PBIMONITOR_ServicePrincipalSecret,
        [string]$resource = "https://api.fabric.microsoft.com/"
    )

    write-verbose "getting authentication token"
    $granttype = "client_credentials"
    $tokenuri = "https://login.microsoftonline.com/$($tenantId)/oauth2/token"
    #$appsecret = [System.Web.HttpUtility]::urlencode($appsecret)
    $body = @{
        grant_type    = $granttype
        client_id     = $appid
        client_secret = $appsecret
        resource      = $resource
    }


    $token = invoke-restmethod -uri $tokenuri -method Post -ContentType "application/x-www-form-urlencoded" -body $body
    $accesstoken = $token.access_token
    write-output $accesstoken
}

function Read-FromGraphAPI {
    [CmdletBinding()]
    param
    (
        [string]        $url,
        [string]        $accessToken,
        [string]        $format = "JSON"
    )

    #https://blogs.msdn.microsoft.com/exchangedev/2017/04/07/throttling-coming-to-outlook-api-and-microsoft-graph/

    try {
        $headers = @{
            'Content-Type'  = "application/json"
            'Authorization' = "Bearer $accessToken"
        }

        $result = Invoke-RestMethod -Method Get -Uri $url -Headers $headers

        if ($format -eq "CSV") {
            ConvertFrom-CSV -InputObject $result | Write-Output
        } else {
            Write-Output $result.value
            while ($result.'@odata.nextLink') {
                $result = Invoke-RestMethod -Method Get -Uri $result.'@odata.nextLink' -Headers $headers
                Write-Output $result.value
            }
        }

    } catch [System.Net.WebException] {
        $ex = $_.Exception

        try {
            $statusCode = $ex.Response.StatusCode

            if ($statusCode -eq 429) {
                $message = "429 Throthling Error - Sleeping..."
                Write-Host $message
                Start-Sleep -Seconds 1000
            } else {
                if ($null -ne $ex.Response) {
                    $statusCode = $ex.Response.StatusCode
                    $stream = $ex.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    $errorContent = $reader.ReadToEnd()
                    $message = "$($ex.Message) - '$errorContent'"
                } else {
                    $message = "$($ex.Message) - 'Empty'"
                }
            }

            Write-Error -Exception $ex -Message $message
        } finally {
            if ($reader) { $reader.Dispose() }
            if ($stream) { $stream.Dispose() }
        }
    }
}

function Read-FromTenantAPI {
    [CmdletBinding()]
    param
    (
        [string]$url,
        [string]$accessToken,
        [string]$format = "JSON"
    )

    #https://blogs.msdn.microsoft.com/exchangedev/2017/04/07/throttling-coming-to-outlook-api-and-microsoft-graph/

    try {
        $headers = @{
            'Content-Type'  = "application/json"
            'Authorization' = "Bearer $accessToken"
        }
        $result = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        Write-Output $result
    } catch [System.Net.WebException] {
        $ex = $_.Exception
        try {
            $statusCode = $ex.Response.StatusCode
            if ($statusCode -eq 429) {
                $message = "429 Throthling Error - Sleeping..."

                Write-Host $message
                Start-Sleep -Seconds 1000
            } else {
                if ($null -ne $ex.Response) {
                    $statusCode = $ex.Response.StatusCode
                    $stream = $ex.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    $errorContent = $reader.ReadToEnd()
                    $message = "$($ex.Message) - '$errorContent'"
                } else {
                    $message = "$($ex.Message) - 'Empty'"
                }
            }

            Write-Error -Exception $ex -Message $message
        } finally {
            if ($reader) { $reader.Dispose() }
            if ($stream) { $stream.Dispose() }
        }
    }
}
