class WordDoc {
    [string]$FilePath
    [System.__ComObject]$Document
    [System.__ComObject]$WordApplication

    WordDoc([string]$FilePath) {
        $this.FilePath = $FilePath
        $this.WordApplication = New-Object -ComObject Word.Application
        $this.Document = $this.WordApplication.Documents.Open($FilePath)
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

    [void] SaveAs([string]$NewFilePath) {
        $this.Document.SaveAs([ref]$NewFilePath)
    }

    [void] Close() {
        $this.Document.Close([ref]$false)  # false を指定して変更を保存
        $this.WordApplication.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApplication) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value, [string]$NewFilePath) {
        $this.SetCustomProperty($PropertyName, $Value)
        $this.SaveAs($NewFilePath)
        $this.Close()
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path $this.FilePath -Force
        Rename-Item -Path $NewFilePath -NewName (Split-Path $this.FilePath -Leaf)
    }
}

# 使用例
$OriginalFilePath = "D:\Github\PS_Script\sample.docx"
$NewFilePath = "D:\Github\PS_Script\sample_temp.docx"
$docClass = [WordDoc]::new($OriginalFilePath)
$docClass.SetCustomPropertyAndSaveAs("CustomProperty2", "Value2", $NewFilePath)
