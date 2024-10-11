# Wordアプリケーションを開始
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# 指定されたWordファイルを開く
$doc = $word.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx")

# ドキュメントプロパティを取得
$properties = $doc.BuiltInDocumentProperties

# 新しいセクションを追加
$range = $doc.Content
$range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
$range.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)
$newSection = $doc.Sections.Item($doc.Sections.Count)

# プロパティの値を保存する配列を作成
$values = @()

# プロパティが存在するか確認し、存在しない場合はデフォルト値を設定
try {
    $creationDate = $properties.Item("Creation Date").Value
    Write-Host "Get: Creation Date = $creationDate"
} catch {
    $creationDate = "N/A"
    Write-Host "Get: Creation Date not found, setting to N/A"
}

try {
    $author = $properties.Item("Author").Value
    Write-Host "Get: Author = $author"
} catch {
    $author = "N/A"
    Write-Host "Get: Author not found, setting to N/A"
}

try {
    $lastAuthor = $properties.Item("Last Author").Value
    Write-Host "Get: Last Author = $lastAuthor"
} catch {
    $lastAuthor = "N/A"
    Write-Host "Get: Last Author not found, setting to N/A"
}

$values += ,("Creation Date", $creationDate)
$values += ,("Author", $author)
$values += ,("Last Edited By", $lastAuthor)

# カスタムプロパティを設定する関数
function Set-OfficeDocCustomProperty {
    [OutputType([boolean])]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [string] $Value,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try
    {
        $customProperties = $Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName,$false, 4, $Value
        try
        {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod,$null,$customProperties,$arrayArgs) | out-null
            Write-Host "Set: $PropertyName = $Value"
        } 
        catch [system.exception] 
        {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            Write-Host "Set: $PropertyName = $Value (updated)"
        }
        return $true
    } 
    catch
    {
        Write-Host "Set: Failed to set $PropertyName"
        return $false
    }
}

# カスタムプロパティを読み取る関数
function Get-OfficeDocCustomProperty {
    [OutputType([string], $null)]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try {
        $comObject = $Document.CustomDocumentProperties($PropertyName)
        $binding = "System.Reflection.BindingFlags" -as [type]
        $val = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$comObject,$null)
        Write-Host "Get: $PropertyName = $val"
        return $val
    } catch {
        Write-Host "Get: $PropertyName not found"
        return $null
    }
}

# カスタムプロパティを設定
$customPropertyName = "Project"
$customPropertyValue = "FA"
Set-OfficeDocCustomProperty -PropertyName $customPropertyName -Value $customPropertyValue -Document $doc

# カスタムプロパティを読み取る
$customPropertyValueRead = Get-OfficeDocCustomProperty -PropertyName $customPropertyName -Document $doc
$values += ,($customPropertyName, $customPropertyValueRead)

# プロパティの値を新しいセクションに追加
$range = $newSection.Range
foreach ($value in $values) {
    $range.InsertAfter("$($value[0]): $($value[1])`n")
}

# Wordファイルを保存して閉じる
$doc.Save()
$doc.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Ready!" -ForegroundColor Green