function Set-OfficeDocCustomPropertyAndSave {
    [OutputType([boolean])]
    Param (
        [Parameter(Mandatory = $true)]
        [string] $PropertyName,
        [Parameter(Mandatory = $true)]
        [string] $Value,
        [Parameter(Mandatory = $true)]
        [System.__ComObject] $Document,
        [Parameter(Mandatory = $true)]
        [System.__ComObject] $WordApplication
    )
    try {
        $customProperties = $Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName, $false, 4, $Value
        try {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        } catch [system.exception] {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }
        $Document.Save()
        $Document.Close([ref]$false)
        $WordApplication.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApplication) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        return $true
    } catch {
        return $false
    }
}

# 使用例
$word = New-Object -ComObject Word.Application
$docFunction = $word.Documents.Open("D:\Github\PS_Script\sample.docx")
Set-OfficeDocCustomPropertyAndSave -PropertyName "CustomProperty" -Value "Value2" -Document $docFunction -WordApplication $word
