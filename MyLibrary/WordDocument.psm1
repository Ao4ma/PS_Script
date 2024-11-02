# MyLibrary/WordDocument.psm1
class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [System.__ComObject]$Document
    [System.__ComObject]$WordApp

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        Write-Host "Initializing WordDocument..."
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        Write-Host "Creating Word Application COM object..."
        $this.WordApp = New-Object -ComObject Word.Application
        $this.WordApp.DisplayAlerts = 0  # wdAlertsNone
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        Write-Host "Opening document: $docPath"
        $this.Document = $this.WordApp.Documents.Open($docPath)
        Write-Host "WordDocument initialized successfully."
    }
}