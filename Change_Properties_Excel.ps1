# Excelアプリケーションを開始
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# 指定されたExcelファイルを開く
$workbook = $excel.Workbooks.Open("C:\Users\y0927\Documents\GitHub\PS_Script\新規図面発行通知書_図面一覧表.xls")

# ドキュメントプロパティを取得
$properties = $workbook.BuiltinDocumentProperties

# 新しいシートを追加
$newWorksheet = $workbook.Sheets.Add()
$newWorksheet.Name = "Metadata"

# プロパティの値を保存する配列を作成
$values = @()

# プロパティが存在するか確認し、存在しない場合はデフォルト値を設定
try {
    $creationDate = $properties.Item("Creation Date").Value
} catch {
    $creationDate = "N/A"
}

try {
    $author = $properties.Item("Author").Value
} catch {
    $author = "N/A"
}

try {
    $lastAuthor = $properties.Item("Last Author").Value
} catch {
    $lastAuthor = "N/A"
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
        } 
        catch [system.exception] 
        {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }
        return $true
    } 
    catch
    {
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
        return $val
    } catch {
        return $null
    }
}

# カスタムプロパティを設定
$customPropertyName = "Project"
$customPropertyValue = "FA"
Set-OfficeDocCustomProperty -PropertyName $customPropertyName -Value $customPropertyValue -Document $workbook

# カスタムプロパティを読み取る
$customPropertyValueRead = Get-OfficeDocCustomProperty -PropertyName $customPropertyName -Document $workbook
$values += ,($customPropertyName, $customPropertyValueRead)

# プロパティの値を新しいシートの範囲に設定
for ($i = 0; $i -lt $values.Length; $i++) {
    $newWorksheet.Cells.Item($i + 1, 1).Value2 = $values[$i][0]
    $newWorksheet.Cells.Item($i + 1, 2).Value2 = $values[$i][1]
}

# Excelファイルを保存して閉じる
$workbook.Save()
$workbook.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Ready!" -ForegroundColor Green


# メソッド
# addCustomProperty(key, value)	
# 新しいカスタム プロパティを作成、または既存のカスタム プロパティを設定します。

# deleteAllCustomProperties()	
# このコレクション内のすべてのカスタム プロパティを削除します。

# getAuthor()	
# ブックの作成者。

# getCategory()	
# ブックのカテゴリ。

# getComments()	
# ブックのコメント。

# getCompany()	
# ブックの会社。

# getCreationDate()	
# ブックの作成日を取得します。

# getCustom()	
# ブックのカスタム プロパティのコレクションを取得します。

# getCustomProperty(key)	
# キーを使用してカスタム プロパティ オブジェクトを取得します。大文字と小文字は区別されません。 カスタム プロパティが存在しない場合、このメソッドは を返します undefined。

# getKeywords()	
# ブックのキーワード。

# getLastAuthor()	
# ブックの最後の作成者を取得します。

# getManager()	
# ブックのマネージャー。

# getRevisionNumber()	
# ブックのリビジョン番号を取得します。

# getSubject()	
# ブックの件名。

# getTitle()	
# ブックのタイトル。

# setAuthor(author)	
# ブックの作成者。

# setCategory(category)	
# ブックのカテゴリ。

# setComments(comments)	
# ブックのコメント。

# setCompany(company)	
# ブックの会社。

# setKeywords(keywords)	
# ブックのキーワード。

# setManager(manager)	
# ブックのマネージャー。

# setRevisionNumber(revisionNumber)	
# ブックのリビジョン番号を取得します。

# setSubject(subject)	
# ブックの件名。

# setTitle(title)	
# ブックのタイトル。