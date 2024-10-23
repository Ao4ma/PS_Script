function Set-DocumentProperty {
    param (
        [string]$Path,          # ドキュメントのパス
        [string]$PropertyName,  # 設定するプロパティ名
        [string]$PropertyValue  # 設定するプロパティ値
    )

    try {
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($Path)

        $binding = "System.Reflection.BindingFlags" -as [type]
        [ref]$SaveOption = "Microsoft.Office.Interop.Word.WdSaveOptions" -as [type]

        $propertySet = $false

        # ビルトインプロパティを設定
        try {
            $Properties = $doc.BuiltInDocumentProperties
            $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $PropertyName)
            [System.__ComObject].InvokeMember("Value", $binding::SetProperty, $null, $pn, $PropertyValue)
            $propertySet = $true
        } catch [System.Exception] {
            Write-Host -ForegroundColor Blue "Builtin property '$PropertyName' not found or cannot be set."
        }

        # カスタムプロパティを設定
        if (-not $propertySet) {
            try {
                $Properties = $doc.CustomDocumentProperties
                $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $PropertyName)
                [System.__ComObject].InvokeMember("Value", $binding::SetProperty, $null, $pn, $PropertyValue)
                $propertySet = $true
            } catch [System.Exception] {
                Write-Host -ForegroundColor Blue "Custom property '$PropertyName' not found or cannot be set."
            }

            # カスタムプロパティが存在しない場合、新規作成
            if (-not $propertySet) {
                try {
                    $Properties.Add($PropertyName, $false, 4, $PropertyValue)  # 4 corresponds to msoPropertyTypeString
                    Write-Host -ForegroundColor Green "Custom property '$PropertyName' created and set to '$PropertyValue'."
                    $propertySet = $true
                } catch [System.Exception] {
                    Write-Host -ForegroundColor Red "Failed to create and set custom property '$PropertyName'."
                }
            }
        }

        # ドキュメントを保存して閉じる
        $doc.Save()
        $doc.Close([ref]$SaveOption::wdSaveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Properties) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        Remove-Variable -Name doc, Properties

        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Remove-Variable -Name word
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()

        if ($propertySet) {
            Write-Host "Property '$PropertyName' set to '$PropertyValue'." -ForegroundColor Green
            return $true
        } else {
            Write-Host "Failed to set property '$PropertyName'." -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        return $false
    }
}

# 関数の呼び出し例
$path = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx"
$propertyName = "batter"
$propertyValue = "近藤さんq"

$result = Set-DocumentProperty -Path $path -PropertyName $propertyName -PropertyValue $propertyValue

if ($result) {
    Write-Host "Property set successfully." -ForegroundColor Green
} else {
    Write-Host "Failed to set property." -ForegroundColor Red
}