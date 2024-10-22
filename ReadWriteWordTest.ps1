# 書き込み例
function Write-DocumentProperty {
    param (
        [string]$filePath,
        [string]$propertyName,
        [string]$propertyValue,
        [switch]$isNewCustomProperty
    )

    # PCクラスのインスタンスを作成
    $pc = New-Object -TypeName PSObject -Property @{
        IsLibraryConfigured = $true
    }

    # Wordクラスのインスタンスを作成
    $word = [Word]::new($filePath, $pc)

    if ($null -eq $word.Document) {
        Write-Error "ドキュメントを開くことができませんでした。ファイルパスを確認してください: $filePath"
        return
    }

    # ドキュメントプロパティを設定
    $word.SetDocumentProperty($propertyName, $propertyValue, $isNewCustomProperty)

    # ドキュメントを閉じる
    $word.Close()
}


# 読み込み例
function Read-DocumentProperties {
    param (
        [string]$filePath
    )

    # PCクラスのインスタンスを作成
    $pc = New-Object -TypeName PSObject -Property @{
        IsLibraryConfigured = $true
    }

    # Wordクラスのインスタンスを作成
    $word = [Word]::new($filePath, $pc)

    if ($null -eq $word.Document) {
        Write-Error "ドキュメントを開くことができませんでした。ファイルパスを確認してください: $filePath"
        return
    }

    # ドキュメントプロパティを取得
    $properties = $word.GetDocumentProperties()

    # プロパティを表示
    foreach ($key in $properties.Keys) {
        Write-Host "$($key): $($properties[$key])"
    }

    # ドキュメントを閉じる
    $word.Close()
}



# デフォルト値を変数として定義
$defaultFilePath = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx"
$defaultPropertyName = "Author"
$defaultPropertyValue = "近藤さんA"

[string]$filePath = $defaultFilePath
[string]$propertyName = $defaultPropertyName
[string]$propertyValue = $defaultPropertyValue
[switch]$isNewCustomProperty

# Word_Class.ps1を読み込む
. "C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\Word_Class.ps1"

# ドキュメントプロパティの書き込み
Write-DocumentProperty -filePath $filePath -propertyName $propertyName -propertyValue $propertyValue -isNewCustomProperty:$isNewCustomProperty

# ドキュメントプロパティの読み込み
Read-DocumentProperties -filePath $filePath
