# PowerShellスクリプト: Set-WordCustomProperties.ps1

param (
    [string]$filePath = "C:\Users\y0927\Documents\GitHub\PS_Script\work\技100-999.docx",
    [string]$approver = "青島",
    [bool]$approvalFlag = $true
)

# Wordアプリケーションを起動
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# ドキュメントを開く
$doc = $word.Documents.Open($filePath)

# カスタムプロパティを設定する関数
function Set-CustomProperty {
    param (
        [object]$doc,
        [string]$propName,
        [object]$propValue
    )

    $properties = $doc.CustomDocumentProperties
    $property = $properties | Where-Object { $_.Name -eq $propName }

    if ($property) {
        # 既存のプロパティを更新
        $property.Value = $propValue
    } else {
        # 新しいプロパティを追加
        $properties.Add($propName, $false, 4, $propValue) # 4はmsoPropertyTypeString
    }
}

# 承認者プロパティを設定
Set-CustomProperty -doc $doc -propName "承認者" -propValue $approver

# 承認フラグプロパティを設定
Set-CustomProperty -doc $doc -propName "承認フラグ" -propValue ([string]$approvalFlag)

# ドキュメントを保存して閉じる
$doc.Save()
$doc.Close()

# Wordアプリケーションを終了
$word.Quit()

Write-Host "カスタムプロパティが設定されました。"