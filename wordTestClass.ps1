class WordDoc {
    [string]$FilePath
    [System.__ComObject]$Document
    [System.__ComObject]$WordApplication

    WordDoc([string]$FilePath) {
        $this.FilePath = $FilePath
        $this.WordApplication = New-Object -ComObject Word.Application
        $this.Document = $this.WordApplication.Documents.Open($FilePath)
    }

    [void] SetCustomPropertyAndSave([string]$PropertyName, [string]$Value) {
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
        $this.Document.Save()
        $this.Document.Close([ref]$false)
        $this.WordApplication.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.Document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.WordApplication) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# 使用例
$docClass = [WordDoc]::new("D:\Github\PS_Script\sample.docx")
$docClass.SetCustomPropertyAndSave("CustomProperty", "Value1")
