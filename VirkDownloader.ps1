# Virk.dk downloader tool

# Made by - Mathias Wrobel
# https://github.com/mathwro/

param (
    [String] $ServerUri = "http://distribution.virk.dk",
    [Int] $size = 50,
    [String] $fileType = "xml",
    [String] $scrollTime = "30m"
)



#Preparing powershell
$tmpFC = $host.PrivateData.ProgressForeGroundColor
$tmpBC = $host.PrivateData.ProgressBackGroundColor
$host.PrivateData.ProgressForeGroundColor = "Black"
$host.PrivateData.ProgressBackGroundColor = "DarkBlue"

Function scrollBody ($scrollID, $scrollTime) {
    $scrollBody = @"
    {
        "scroll" : "$scrollTime",
        "scroll_id": "$scrollID"
    }
"@
    return $scrollBody
}

Function fileDownloader ($response, $filePath) {
    ForEach ($hit in $response.hits.hits) {
        $documentURL = ($hit._source.dokumenter | Where-Object { $_.dokumentMimeType -eq "application/$fileType" }).dokumentUrl
        #Verifying URL is present, otherwise skip
        if ($documentURL) {
            if ($documentURL.Count -gt 1) {
                $documentURL = $documentURL[0]
            }
            $randomTimeID = Get-Date -Format fffff
            $finalPath = $filePath + "/" + $hit._source.cvrNummer + "_" + $hit._index + "id" + $randomTimeID + ".xml"
            (Invoke-WebRequest `
                    -Method GET `
                    -Uri $documentURL).Content `
            | Out-File -FilePath $finalPath
        }
    }
}

#Get dates
$sDate = Read-Host "Starting date (Format: YYYY-MM-DD)"
if (($sDate -Match "\d{4}\-\d{2}\-\d{2}") -eq $false) {
    do {
        $sDate = Read-Host "Wrong format. Try again"
    } until (($sDate -Match "\d{4}\-\d{2}\-\d{2}") -eq $true)
}
    
$eDate = Read-Host "Ending date (Format: YYYY-MM-DD)"
if (($eDate -Match "\d{4}\-\d{2}\-\d{2}") -eq $false) {
    do {
        $eDate = Read-Host "Wrong format. Try again"
    } until (($eDate -Match "\d{4}\-\d{2}\-\d{2}") -eq $true)
}

#Setting up parameters
$filePath = Read-Host "Enter file path (Leave empty to save in the script folder)"
if ( -not ($filePath)) {
    if ((Test-Path -Path "$PSScriptRoot\downloads") -eq $false) {
        New-Item -ItemType Directory -Name "downloads" -Path "$PSScriptRoot"
    }
    $filePath = "$PSScriptRoot\downloads"
}

#Setting up scroll uri
$initUri = $ServerUri + "/offentliggoerelser/_search" + "?scroll=$scrollTime"
$scrollUri = $ServerUri + "/_search/scroll/"

#Setting date strings
$startingDate = $sDate + "T00:00:00.001Z"
$endingDate = $eDate + "T23:59:59.505Z"

$body = @{
    "size" = $size
    "query" = @{
        "bool" = @{
            "filter" = @(
                @{
                    "term" = @{
                        "dokumenter.dokumentMimeType" = "application"
                    }
                }
                @{
                    "term" = @{
                        "dokumenter.dokumentMimeType" = "$fileType"
                    }
                }
                @{
                    "range" = @{
                        "offentliggoerelsesTidspunkt" = @{
                            "gte" = "$startingDate"
                            "lte" = "$endingDate"
                        }
                    }
                }
            )
        }
    }
} | ConvertTo-Json -Depth 6


#Invoking API to get total hits
$initalParams = @{
    'Uri'           = $initUri
    'Method'        = 'POST'
    'ContentType'   = 'application/json'
    'Body'          = $body
}
$response = (Invoke-RestMethod @initalParams)


#Setting up additional parameters
$randomTimeID = Get-Date -Format fffff
$folderName = "Dato_" + $sDate + "_" +$eDate + "_Antal_" + $response.hits.total + "_ID_" + $randomTimeID
New-Item -ItemType Directory -Name $folderName -Path $filePath | Out-Null
$filePath = $filePath + "\" + $folderName
Write-Host "`n"


#Running information
Write-Host "Total hits:" $response.hits.total
Write-Host "Max score:" $response.hits.max_score
Write-Host "`n"
Write-Host "We are still working :)"
Write-Host "(Press Ctrl+C to cancel)"

#Setting up for scroll
$scrollID = $response._scroll_id
$scrollAmount = [Math]::Ceiling($response.hits.total / $size)
$singlePercent = 100 / $scrollAmount

fileDownloader `
    -response $response `
    -filePath $filePath

$currentPercent = 0
For ($i = 1; $i -le $scrollAmount; $i++) {
    $currentPercent = $currentPercent + $singlePercent
    $roundedPercent = [math]::Round($currentPercent, 2)
    Write-Progress -Activity "`r$roundedPercent% : Downloading files" -PercentComplete $currentPercent


    $body = @{
        "scroll"    = "$scrollTime"
        "scroll_id" = "$scrollID"
    } | ConvertTo-Json

    $scrollParams = @{
        'Uri'             = $scrollUri
        'Method'          = 'POST'
        'ContentType'     = 'application/json'
        'Body'            = $body
    }
    $response = Invoke-RestMethod @scrollParams

    fileDownloader `
        -response $response `
        -filePath $filePath
}
Write-Host "Job's done"
$host.PrivateData.ProgressForeGroundColor = $tmpFC
$host.PrivateData.ProgressBackGroundColor = $tmpBC