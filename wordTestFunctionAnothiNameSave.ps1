function Set-OfficeDocCustomPropertyAndSaveAs {
    [OutputType([boolean])]
    Param (
        [Parameter(Mandatory = $true)]
        [string] $PropertyName,
        [Parameter(Mandatory = $true)]
        [string] $Value,
        [Parameter(Mandatory = $true)]
        [string] $OriginalFilePath,
        [Parameter(Mandatory = $true)]
        [string] $NewFilePath
    )
    try {
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($OriginalFilePath)

        $customProperties = $doc.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName, $false, 4, $Value
        try {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        } catch [system.exception] {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }

        # 別名で保存
        $doc.SaveAs([ref]$NewFilePath)
        $doc.Close()
        $word.Quit()

        # COMオブジェクトの解放
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        # 元のファイルを削除し、別名で保存したファイルをリネーム
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path $OriginalFilePath -Force
        Rename-Item -Path $NewFilePath -NewName (Split-Path $OriginalFilePath -Leaf)

        return $true
    } catch {
        Write-Error $_.Exception.Message
        return $false
    }
}

# 使用例
$OriginalFilePath = "D:\Github\PS_Script\sample.docx"
$NewFilePath = "D:\Github\PS_Script\sample_temp.docx"
Set-OfficeDocCustomPropertyAndSaveAs -PropertyName "CustomProperty" -Value "Value1" -OriginalFilePath $OriginalFilePath -NewFilePath $NewFilePath
