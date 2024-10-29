class WordDocument {
    [void] Set_Custom_Property([string]$PropertyName, [string]$Value, [System.__ComObject]$Document) {
        Write-Host "IN: Set_Custom_Property"
        try {
            $customProperties = $Document.CustomDocumentProperties
            $binding = "System.Reflection.BindingFlags" -as [type]
            [array]$arrayArgs = $PropertyName, $false, 4, $Value
            $propertyExists = $false
            $propertyObject = $null
            try {
                $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
                $propertyExists = $true
            } catch {
                $propertyExists = $false
            }
            if ($propertyExists) {
                $propertyObject.Value = $Value
            } else {
                [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) 
            }
            Write-Host "Property '$PropertyName' set to '$Value'."
            
            # カスタム属性を設定後すぐに確認
            $propValue = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName).Value
            Write-Host "Immediately After Setting - Property: $PropertyName, Value: $propValue"
            
            # 保存前にカスタム属性を確認
            foreach ($prop in $customProperties) {
                Write-Host "Before Save - Property: $($prop.Name), Value: $($prop.Value)"
            }
            
            # ドキュメントを別名で保存
            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newDocPath = $Document.FullName -replace '\.docx$', "_$timestamp.docx"
            $Document.SaveAs([ref]$newDocPath)
            
            # 保存後にカスタム属性を確認
            $customPropertiesAfterSave = $Document.CustomDocumentProperties
            foreach ($prop in $customPropertiesAfterSave) {
                Write-Host "After Save - Property: $($prop.Name), Value: $($prop.Value)"
            }
        } catch {
            Write-Host "Failed to set property '$PropertyName': $_" -ForegroundColor Red
        }
        Write-Host "OUT: Set_Custom_Property"
    }
}

# クラスのインスタンスを作成
$wordDoc = [WordDocument]::new()

# メソッドの呼び出し例
$docPath = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\技100-999.docx"
if (-Not (Test-Path $docPath)) {
    Write-Host "File not found: $docPath" -ForegroundColor Red
    exit
}
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($docPath)
$wordDoc.Set_Custom_Property("CustomPropT", "CustomValueT", $doc)
try {
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $newDocPath = $docPath -replace '\.docx$', "_$timestamp.docx"
    $doc.SaveAs([ref]$newDocPath)
    $doc.Close()
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) 
    Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) 
    Out-Null
    Remove-Variable -Name doc, word
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    
    # 元のファイルを削除して新しいファイルをリネーム
    Remove-Item -Path $docPath
    Rename-Item -Path $newDocPath -NewName $docPath
} catch {
    Write-Host "Failed to save document: $_" -ForegroundColor Red
}
