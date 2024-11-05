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

    [void] SaveAs([string]$NewFilePath) {
        $this.Document.SaveAs([ref]$NewFilePath)
    }

    [void] Close() {
        $this.Document.Close([ref]$false)  # false を指定して変更を保存しない
        $this.WordApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}