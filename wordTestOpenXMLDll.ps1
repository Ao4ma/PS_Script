# 必要な名前空間をインポート
using namespace DocumentFormat.OpenXml.Packaging
using namespace DocumentFormat.OpenXml.CustomProperties

# Open XML SDKのアセンブリを読み込む
$openXmlPath = "D:\Github\PS_Script\lib\net46\DocumentFormat.OpenXml.dll"
Add-Type -Path $openXmlPath

function Add-CustomProperty {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$filePath,
        [Parameter(Mandatory = $true)]
        [string]$propertyName,
        [Parameter(Mandatory = $true)]
        [string]$propertyValue
    )

    try {
        # ドキュメントを開く
        $wordDoc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($filePath, $true)

        # カスタムプロパティパートを取得または作成
        if ($null -eq $wordDoc.CustomFilePropertiesPart) {
            $customPropsPart = $wordDoc.AddCustomFilePropertiesPart()
            $props = New-Object DocumentFormat.OpenXml.CustomProperties.Properties
            $customPropsPart.Properties = $props
        } else {
            $customPropsPart = $wordDoc.CustomFilePropertiesPart
        }

        $props = $customPropsPart.Properties

        # 既存プロパティのチェックと削除
        $existingProp = $props | Where-Object { $_.Name -eq $propertyName }
        if ($existingProp) {
            $props.RemoveChild($existingProp)
        }

        # 新しいカスタムプロパティを追加
        $newProp = New-Object DocumentFormat.OpenXml.CustomProperties.CustomProperty
        $newProp.Name = $propertyName
        $newProp.VTLPWSTR = $propertyValue
        $props.AppendChild($newProp)

        # 保存して閉じる
        $wordDoc.Save()
        $wordDoc.Close()
        return $true
    } catch {
        Write-Error $_.Exception.Message
        return $false
    }
}

# 使用例
$filePath = "D:\Github\PS_Script\sample.docx"
$propertyName = "CustomProperty"
$propertyValue = "Value1"

Add-CustomProperty -filePath $filePath -propertyName $propertyName -propertyValue $propertyValue
