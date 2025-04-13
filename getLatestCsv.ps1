$documentId = "1Sj8j0v1lWyaIkP2ICSaVUE1vgpAIHR_c88zvuMlmXb4"
$pageIds = @('1184276413', '1894219008', '48450441', '83947440', '1946812712')



Push-Location $PSScriptRoot
$outputFolder = "config"
Foreach ($pageId in $pageIds) {
    $URL = "https://docs.google.com/spreadsheets/d/$documentId/export?format=csv&gid=$pageId"

    $outputFile = "$outputFolder\$pageId.csv"

    try {
        Invoke-WebRequest -Uri $URL -OutFile $outputFile
        Write-Host "Downloaded: $outputFile"
    } catch {
        Write-Host "Failed to download CSV for Page ID: $pageId" -ForegroundColor Red
    }
}

Write-Host "Download process complete." -ForegroundColor Green
Pop-Location
