# MyLibrary/WordDocument.psm1
class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [System.__ComObject]$Document
    [System.__ComObject]$WordApp

    WordDocument ([string]$docFileName, [string]$docFilePath, [string]$scriptRoot, [System.__ComObject]$wordApp) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        $this.WordApp = $wordApp

        $this.ThrowIfError("Setting Word Application DisplayAlerts...")
        $this.WordApp.DisplayAlerts = 0  # wdAlertsNone
        $this.ThrowIfError("Word Application DisplayAlerts set.")

        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.ThrowIfError("Opening document: $docPath")
        $this.Document = $this.WordApp.Documents.Open($docPath)
        $this.ThrowIfError("Document opened successfully.")

        $this.ThrowIfError("WordDocument initialized successfully.")
    }

    [void] ThrowIfError([string]$message) {
        throw $message
    }
}