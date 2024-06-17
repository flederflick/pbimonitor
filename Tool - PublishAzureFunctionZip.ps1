
$ErrorActionPreference = "Stop"

$currentPath = (Split-Path $MyInvocation.MyCommand.Definition -Parent)
$azureFunctionDevFolder = "$currentPath\AzureFunction"
$publishFolder = "$currentPath"
$publishFolderTemp = "$publishFolder\temp"

#ensure folders
New-Item -ItemType Directory -Path $publishFolderTemp  -Force -ErrorAction SilentlyContinue | Out-Null
New-Item -ItemType Directory -Path $publishFolderTemp\Scripts -Force -ErrorAction SilentlyContinue | Out-Null

#copy azure function
Get-ChildItem $azureFunctionDevFolder -Exclude @(".*","local.*") | Copy-Item -Destination $publishFolderTemp -Force -Recurse

$scripts = @(Get-ChildItem -File -Path "$currentPath\Fetch*.ps1") + @(Get-ChildItem -File -Path "$currentPath\Fetch*.psm1")
$scripts | Copy-Item -Destination "$publishFolderTemp\Scripts" -Force

#zip

Compress-Archive -Path "$publishFolderTemp\*" -CompressionLevel Optimal -DestinationPath "$publishFolder\AzureFunction.zip" -Force

Remove-Item $publishFolderTemp -Force -Recurse

Write-Host "Next Steps:"
Write-Host "- Create an Azure Function with PowerShell"
Write-Host "- Using Kudo ZIP Deploy, deploy the AzureFunction.zip"
Write-Host "- Configure the Configuration Settings on Azure Function"
Write-Host "- Run the function"



$azAccount = Get-InstalledModule az.accounts -ErrorAction SilentlyContinue
$AzWebapp = Get-InstalledModule az.websites -ErrorAction SilentlyContinue

if($azAccount -eq $null -or $AzWebapp -eq $null){
    Write-Host "Please install the following modules:"
    Write-Host "Install-Module -Name Az.Accounts -AllowClobber -Scope CurrentUser"
    Write-Host "Install-Module -Name Az.Websites -AllowClobber -Scope CurrentUser"
    exit
}

Set-AzContext -Subscription hannl-pg-ict-sub
Publish-AzWebapp -ResourceGroupName phannlcbpbimonrg01 -Name phannlcbpbimonfa01 -ArchivePath "$currentPath\AzureFunction.zip" -Type zip -Clean -Restart

