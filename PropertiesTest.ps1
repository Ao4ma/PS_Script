$path = "D:\GDrive\PS_Script\技100-999.docx"
$AryProperties = "Title", "Author", "Keywords", "Number of words", "Number of pages"
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($path)

$binding = "System.Reflection.BindingFlags" -as [type]
[ref]$SaveOption = "Microsoft.Office.Interop.Word.WdSaveOptions" -as [type]

$BuiltinProperties = $doc.BuiltInDocumentProperties
$objHash = @{"Path" = $doc.FullName}

foreach ($p in $AryProperties) {
    try {
        $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $BuiltinProperties, $p)
        $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
        $objHash.Add($p, $value)
    } catch [System.Exception] {
        Write-Host -ForegroundColor Blue "Value not found for $p)"
    }
}

$docProperties = New-Object PSObject -Property $objHash
$docProperties

# ドキュメントを保存せずに閉じる
$doc.Close([ref]$SaveOption::wdDoNotSaveChanges)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
Remove-Variable -Name doc, BuiltinProperties

$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable -Name word
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Ready!" -ForegroundColor Green