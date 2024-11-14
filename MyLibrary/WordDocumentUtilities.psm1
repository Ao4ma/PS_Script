<#
function Write_ToFile {
    param (
        [string]$FilePath,
        [array]$Content
    )
    if ($Content.Count -eq 0) {
        Write-Host "No content found. Deleting previous output file if it exists."
        if (Test-Path $FilePath) {
            Remove-Item $FilePath
        }
    } else {
        $Content | Out-File -FilePath $FilePath -Encoding UTF8
    }
}
#>
function SaveDocumentWithBackup {
    param (
        [WordDocument]$wordDoc
    )
    Write-Host "IN: SaveDocumentWithBackup"
    $backupFilePath = Join-Path -Path $wordDoc.DocFilePath -ChildPath "backup_$($wordDoc.DocFileName)"
    $wordDoc.SaveAs($backupFilePath)
    Write-Host "OUT: SaveDocumentWithBackup"
}