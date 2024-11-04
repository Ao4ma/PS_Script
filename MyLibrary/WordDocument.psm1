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

        # WordのDisplayAlertsを無効にする
        $this.WordApp.DisplayAlerts = 0  # wdAlertsNone

        # ドキュメントを開く
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.Document = $this.WordApp.Documents.Open($docPath)
    }

    [void] SaveAs([string]$NewFilePath) {
        $this.Document.SaveAs([ref]$NewFilePath)
    }

    [void] Close() {
        $this.Document.Close([ref]$false)
        $this.WordApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    [void] SetCustomProperty([string]$PropertyName, [string]$Value) {
        $customProperties = $this.Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName, $false, 4, $Value
        try {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        } catch [system.exception] {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }
    }

    [void] FillSignatures() {
        $table = $this.Document.Tables.Item(1) # 1番目の表を取得
        # カスタムプロパティから担当者と承認者の名前と日付を取得
        $担当者 = [System.__ComObject].InvokeMember("Item", "GetProperty", $null, $this.Document.CustomDocumentProperties, "担当者").Value
        $承認者 = [System.__ComObject].InvokeMember("Item", "GetProperty", $null, $this.Document.CustomDocumentProperties, "承認者").Value
        $today = Get-Date -Format "yyyy年MM月dd日"

        # 担当者のサイン欄
        $cell1 = $table.Cell(2, 1) # 2行1列目
        $cell1.Range.Text = "$担当者`n$today"
        $cell1.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter

        # 承認者のサイン欄
        $cell2 = $table.Cell(2, 2) # 2行2列目
        $cell2.Range.Text = "$承認者`n$today"
        $cell2.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
    }
}
