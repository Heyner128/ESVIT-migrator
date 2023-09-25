Get-ChildItem -Recurse | Unblock-File
Import-Module (Get-ChildItem -Recurse -Filter "*.psd1").FullName -DisableNameChecking

Write-Banner "ESVIT migrator"
Write-Host "Created by heyner.cuevas@skf.com for SKF France based on https://github.com/Zerg00s/FlowPowerAppsMigrator" -ForegroundColor Yellow
Write-Host "Please don't use it without the authorisation of the author" -ForegroundColor Yellow
Write-Host

$SOURCE_SITE_URL="https://skfgroup.sharepoint.com/sites/O365-EMEAITOTSecurity"
$TARGET_SITE_URL=Read-Host "Target site URL"

if($SOURCE_SITE_URL -eq $TARGET_SITE_URL) {
    throw "The target site cannot be same as source site"
} 

try {
    Write-Host
    Write-Host "Connecting to $SOURCE_SITE_URL" -ForegroundColor Magenta
    $sourceSiteConnection = Connect-PnPOnline -Url $SOURCE_SITE_URL -UseWebLogin -WarningAction Ignore -ReturnConnection
    Write-Host "Connected to $SOURCE_SITE_URL" -ForegroundColor Magenta
    Write-Host
} catch {
    throw "Source site connection error " + $error[0].Exception.Message
}

try {
    Write-Host
    Write-Host "Connecting to $TARGET_SITE_URL" -ForegroundColor Magenta
    $targetSiteConnection = Connect-PnPOnline -Url $TARGET_SITE_URL -UseWebLogin -WarningAction Ignore -ReturnConnection
    Write-Host "Connected to $TARGET_SITE_URL" -ForegroundColor Magenta
    Write-Host
} catch {
    throw "Target site connection error " + $error[0].Exception.Message
}


$MAIN_LIST_TITLE="LIVR"
$SITES_LIST_TITLE="SKF sites"


try {
    $sourceLists = Get-PnPList -Connection $sourceSiteConnection
    $targetLists = Get-PnPList -Connection $targetSiteConnection
}
catch {
    if ($error[0].Exception.Message -match "(403)" -or $error[0].Exception.Message -match "unauthorized") {
        throw " Make sure you are member at the source site $SOURCE_SITE_URL and target site $TARGET_SITE_URL"
    }
    else {
        throw $error[0].Exception.Message
    }
}

function ContainsLists([Microsoft.Sharepoint.Client.List[]] $lists) {
    return ($lists | Where-Object {$_.Title -eq $MAIN_LIST_TITLE -or $_.Title -eq $SITES_LIST_TITLE}).Count -eq 2
}
         
if(!(ContainsLists($sourceLists))) {
    Write-Host "$SOURCE_SITE_URL doesn't contain the '$MAIN_LIST_TITLE' or the '$SITES_LIST_TITLE' lists" -ForegroundColor Red
    throw "Required lists not found"
}

if(!(ContainsLists($targetLists))) {
    Write-Host "$TARGET_SITE_URL doesn't contain the '$MAIN_LIST_TITLE' or the '$SITES_LIST_TITLE' lists" -ForegroundColor Red
    throw "Required lists not found"
}

function Get-MainListId([Microsoft.Sharepoint.Client.List[]] $lists) {
    return ($lists | Where-Object {$_.Title -eq $MAIN_LIST_TITLE})[0].Id
}

function Get-SitesListId([Microsoft.Sharepoint.Client.List[]] $lists) {
    return ($lists | Where-Object {$_.Title -eq $SITES_LIST_TITLE})[0].Id
}

$sourceMainListId = Get-MainListId($sourceLists)

$targetMainListId = Get-MainListId($targetLists)

$sourceSitesListId = Get-SitesListId($sourceLists)

$targetSitesListId = Get-SitesListId($targetLists)

$CurrentDir = $PSScriptRoot

Remove-item $CurrentDir\package-temp -Force -Recurse -ErrorAction SilentlyContinue

Copy-Item -LiteralPath $CurrentDir\package -Destination $CurrentDir\package-temp -Force -Recurse

Get-ChildItem -LiteralPath $CurrentDir\package-temp -Recurse -Attributes !Directory -Filter "*.json" | ForEach-Object{
    (Get-Content -LiteralPath $_.FullName) -replace $sourceMainListId, $targetMainListId | Set-Content -LiteralPath $_.FullName
    (Get-Content -LiteralPath $_.FullName) -replace $sourceSitesListId, $targetSitesListId | Set-Content -LiteralPath $_.FullName
    (Get-Content -LiteralPath $_.FullName) -replace $SOURCE_SITE_URL, $TARGET_SITE_URL | Set-Content -LiteralPath $_.FullName
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    $msapp = (Get-ChildItem -Path $CurrentDir\package-temp -Recurse -Filter *.msapp)[0]
} catch {
    throw "Incorrect extract package"
}



$msappExtractDirectory = $msapp.Directory.FullName + "\" + $msapp.BaseName


[System.IO.Compression.ZipFile]::ExtractToDirectory($msapp.FullName, $msappExtractDirectory)


Get-ChildItem -Path $msappExtractDirectory -Recurse -File -Filter "*.json" | ForEach-Object{
    (Get-Content -LiteralPath $_.FullName) -replace $sourceMainListId, $targetMainListId | Set-Content -LiteralPath $_.FullName
    (Get-Content -LiteralPath $_.FullName) -replace $sourceSitesListId, $targetSitesListId | Set-Content -LiteralPath $_.FullName
    (Get-Content -LiteralPath $_.FullName) -replace $SOURCE_SITE_URL, $TARGET_SITE_URL | Set-Content -LiteralPath $_.FullName
}

Remove-Item $msapp.FullName -Force

[System.IO.Compression.ZipFile]::CreateFromDirectory($msappExtractDirectory, $msapp.FullName, [System.IO.Compression.CompressionLevel]::Optimal, $false)

Remove-Item $msappExtractDirectory -Force -Recurse -ErrorAction SilentlyContinue

function Save-File([string] $initialDirectory){

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    
    $exportDate = Get-Date -Format "yyyyMMdd"

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.FileName = "ESVIT_export_" + $exportDate + ".zip"
    $OpenFileDialog.filter = "Zip files (*.zip)| *.zip"
    $OpenFileDialog.ShowDialog() |  Out-Null

    return $OpenFileDialog.filename
}

Write-Host "Please choose the file location" -ForegroundColor Magenta

$SaveFileLocation = Save-File $env:USERPROFILE

if($SaveFileLocation -match "\\") {

    Remove-Item $SaveFileLocation -Recurse -Force -ErrorAction SilentlyContinue

    [System.IO.Compression.ZipFile]::CreateFromDirectory("$CurrentDir\package-temp", $SavefileLocation, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    Remove-Item $CurrentDir\package-temp -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "ESVIT package succesfull exported to $SaveFileLocation" -ForegroundColor Green
} else {
    Write-Host "[Warning] No destination folder was choosen" -ForegroundColor Yellow
}


    











