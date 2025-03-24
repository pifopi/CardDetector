$DOCUMENT_ID = "1Qg3NHmUFMPQ1QL9XoViZMtpw1kJ0NWk80SXACUhY0dE"
$PAGE_IDS = @('1184276413', '1894219008', '48450441', '83947440')



$OUTPUT_FOLDER = "config"
Foreach ($PAGE_ID in $PAGE_IDS) {
    $CSV_URL = "https://docs.google.com/spreadsheets/d/$DOCUMENT_ID/export?format=csv&gid=$PAGE_ID"

    $OUTPUT_FILE = "$OUTPUT_FOLDER\$PAGE_ID.csv"

    try {
        Invoke-WebRequest -Uri $CSV_URL -OutFile $OUTPUT_FILE
        Write-Host "Downloaded: $OUTPUT_FILE"
    } catch {
        Write-Host "Failed to download CSV for Page ID: $PAGE_ID" -ForegroundColor Red
    }
}

Write-Host "Download process complete." -ForegroundColor Green
