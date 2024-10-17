# Word_Class.ps1を読み込む
. "C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\Word_Class.ps1"

# パラメータの定義
param (
    [string]$filePath = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx",
    [string]$propertyName = "Author",
    [string]$propertyValue = "近藤さん"
)

# 書き込み例
function Write-DocumentProperty {
    param (
        [string]$filePath,
        [string]$propertyName,
        [string]$propertyValue
    )

    # PCクラスのインスタンスを作成
    $pc = New-Object -TypeName PSObject -Property @{
        IsLibraryConfigured = $true
    }

    # Wordクラスのインスタンスを作成
    $word = [Word]::new($filePath, $pc)

    # ドキュメントプロパティを設定
    $word.SetDocumentProperty($propertyName, $propertyValue)

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

    # ドキュメントプロパティを取得
    $properties = $word.GetDocumentProperties()

    # プロパティを表示
    foreach ($key in $properties.Keys) {
        Write-Host "$($key): $($properties[$key])"
    }

    # ドキュメントを閉じる
    $word.Close()
}

# ドキュメントプロパティの書き込み
Write-DocumentProperty -filePath $filePath -propertyName $propertyName -propertyValue $propertyValue

# ドキュメントプロパティの読み込み
Read-DocumentProperties -filePath $filePath