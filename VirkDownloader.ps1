# Virk.dk downloader tool

# Made by - Mathias Wrobel

param (
    [String] $ServerUri = "http://distribution.virk.dk",
    [Int] $size = 3000,
    [String] $fileType = "xml"
)



#Preparing powershell
$tmpFC = $host.PrivateData.ProgressForeGroundColor
$tmpBC = $host.PrivateData.ProgressBackGroundColor
$host.PrivateData.ProgressForeGroundColor = "Black"
$host.PrivateData.ProgressBackGroundColor = "DarkBlue"

Function newBody ($startingDate, $endingDate, $fileType) {
    $jsonBody = @"
    {
        "query": {
            "bool": {
                "must": [
                    {
                        "term": {
                            "dokumenter.dokumentMimeType": "application"
                        }
                    },
                    {
                        "term": {
                            "dokumenter.dokumentMimeType": "$fileType"
                        }
                    }
                ],
                "filter": {
                    "range": {
                            "offentliggoerelsesTidspunkt": {
                                "gt": "$startingDate",
                                "lt": "$endingDate"
                            }
                    }
                }
            }
        }
    }
"@
    return $jsonBody
}

Function scrollBody ($scrollID) {
    $scrollBody = @"
    {
        "scroll" : "1m",
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

#Setting up additional parameters
$timeNow = Get-Date -Format "yyyy-MMdd_HHmm"
New-Item -ItemType Directory -Name $timeNow -Path $filePath | Out-Null
$filePath = $filePath + "\" + $timeNow
Write-Host "`n"

#Setting up scroll uri
$initUri = $ServerUri + "/offentliggoerelser/_search" + "?scroll=1m&size=$size"
$scrollUri = $ServerUri + "/_search/scroll/"

#Setting date strings
$startingDate = $sDate + "T00:00:00.001Z"
$endingDate = $eDate + "T23:59:59.505Z"

$body = (newBody `
        -startingDate $startingDate `
        -endingDate $endingDate `
        -size $size `
    | ConvertFrom-Json)

#Invoking API to get total hits
$response = (Invoke-WebRequest `
        -Method GET `
        -Uri $initUri `
        -ContentType 'application/json' `
        -Body $body)
$response = $response | ConvertFrom-Json

Write-Host "Total hits:" $response.hits.total
Write-Host "Max score:" $response.hits.max_score
Write-Host "`n"

#Setting up for scroll
$scrollID = $response._scroll_id
$scrollAmount = [Math]::Ceiling($response.hits.total / $size)
$singlePercent = 100 / $scrollAmount

fileDownloader `
    -response $response `
    -filePath $filePath

For ($i = 1; $i -le $scrollAmount; $i++) {
    $currentPercent = $currentPercent + $singlePercent
    $roundedPercent = [math]::Round($currentSubPercent, 2)
    Write-Progress -Activity "`r$roundedPercent% : Downloading files" -PercentComplete $currentPercent
    
    $body = (scrollBody `
            -scrollID $scrollID `
        | ConvertFrom-Json)

    $scrollParams = @{
        'Uri'             = $scrollUri
        'Method'          = 'GET'
        'ContentType'     = 'applications/json'
        'Body'            = @{
            'scroll'    = '2m'
            'scroll_id' = $scrollId
        }
        'UseBasicParsing' = $true
    }
    $response = Invoke-WebRequest @scrollParams | ConvertFrom-Json

    fileDownloader `
        -reponse $response `
        -filePath $filePath
}
$host.PrivateData.ProgressForeGroundColor = $tmpFC
$host.PrivateData.ProgressBackGroundColor = $tmpBC