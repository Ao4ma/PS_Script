# MyLibrary/WordDocument.psm1
class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [System.__ComObject]$Document
    [System.__ComObject]$WordApp

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        $this.WordApp = New-Object -ComObject Word.Application
        $this.WordApp.DisplayAlerts = 0  # wdAlertsNone
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.Document = $this.WordApp.Documents.Open($docPath)
    }

    # メソッドのインポート
    using module "$PSScriptRoot\WordDocumentProperties.psm1"
    using module "$PSScriptRoot\WordDocumentUtilities.psm1"
    using module "$PSScriptRoot\WordDocumentSignatures.psm1"
    using module "$PSScriptRoot\WordDocumentChecks.psm1"
}