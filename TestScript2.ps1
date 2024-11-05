class WordDocument {
    [string]$FilePath
    [System.__ComObject]$Document
    [System.__ComObject]$WordApplication

    WordDocument([string]$FilePath) {
        $this.FilePath = $FilePath
        $this.WordApplication = New-Object -ComObject Word.Application
        $this.Document = $this.WordApplication.Documents.Open($FilePath)
    }

    # ドキュメントを閉じるメソッド
    [void] CloseDocument() {
        if ($null -ne $this.Document) {
            $this.Document.Close([ref]0)  # 0はwdDoNotSaveChangesに相当
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
            Remove-Variable -Name Document
        }
        if ($null -ne $this.WordApplication) {
            $this.WordApplication.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApplication) | Out-Null
            Remove-Variable -Name WordApplication
        }
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
    }

    # ドキュメントを別名で保存するメソッド
    [void] SaveDocumentWithBackup() {
        try {
            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newDocPath = $this.FilePath -replace '\.docx$', "_$timestamp.docx"
            $this.Document.SaveAs([ref]$newDocPath)

            # Close the document and release the COM objects
            $this.CloseDocument()

            # 元のファイルを削除して新しいファイルをリネーム
            Remove-Item -Path $this.FilePath
            Rename-Item -Path $newDocPath -NewName $this.FilePath
        } catch {
            Write-Host "Failed to save document: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # カスタム属性を設定するメソッド
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
}

# ドキュメントのパスとファイル名を設定
$FilePath = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx"

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($FilePath)

# 必要な操作をここに追加
$wordDoc.SetCustomProperty("CustomPropertyName", "CustomValue")

# ドキュメントを別名で保存してからリネーム
$wordDoc.SaveDocumentWithBackup()

# ドキュメントを閉じる
$wordDoc.CloseDocument()