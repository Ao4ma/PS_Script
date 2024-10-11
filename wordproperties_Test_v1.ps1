Param(
    $path = "C:\Users\y0927\Documents\GitHub\PS_Script",
    [array]$include = @("HSG*.docx", "WES*.docx")
)

$AryProperties = "Title", "Author", "Keywords", "Number of words", "Number of pages"

# Wordアプリケーションを開始
$application = New-Object -ComObject Word.Application
$application.Visible = $false

$binding = "System.Reflection.BindingFlags" -as [type]
[ref]$SaveOption = "Microsoft.Office.Interop.Word.WdSaveOptions" -as [type]

# 指定されたパス内のドキュメントを取得
$docs = Get-ChildItem -Path $path -Recurse -Include $include

# ドキュメントごとに処理
foreach ($doc in $docs) {
    $document = $application.Documents.Open($doc.FullName)
    $BuiltinProperties = $document.BuiltInDocumentProperties
    $objHash = @{"Path" = $doc.FullName}

    # ビルトインプロパティを取得
    foreach ($p in $AryProperties) {
        try {
            $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $BuiltinProperties, $p)
            $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
            $objHash.Add($p, $value)
        } catch [System.Exception] {
            Write-Host -ForegroundColor Blue "Value not found for $p"
        }
    }

    # カスタムオブジェクトを作成して表示
    $docProperties = New-Object PSObject -Property $objHash
    $docProperties

    # ドキュメントを保存せずに閉じる
    $document.Close([ref]$SaveOption::wdDoNotSaveChanges)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($BuiltinProperties) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    Remove-Variable -Name document, BuiltinProperties
}

# Wordアプリケーションを終了
$application.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
Remove-Variable -Name application
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Ready!" -ForegroundColor Green