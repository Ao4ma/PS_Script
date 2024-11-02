# MyLibrary/WordDocumentChecks.psm1
# PC環境をチェックするメソッド
[void] Check_PC_Env() {
    Write-Host "IN: Check_PC_Env"
    $osVersion = [System.Environment]::OSVersion.Version
    Write-Host "OS Version: $osVersion"
    Write-Host "OUT: Check_PC_Env"
}

# Wordライブラリをチェックするメソッド
[void] Check_Word_Library() {
    Write-Host "IN: Check_Word_Library"
    try {
        $wordApp = New-Object -ComObject Word.Application
        Write-Host "Microsoft Word is installed."
        $wordApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    } catch {
        Write-Error "Microsoft Word is not installed."
    }
    Write-Host "OUT: Check_Word_Library"
}

# カスタムプロパティをチェックするメソッド
[void] Check_Custom_Property() {
    Write-Host "IN: Check_Custom_Property"
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot check properties.")) {
        Write-Host "OUT: Check_Custom_Property"
        return
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $propCount = [System.__ComObject].InvokeMember("Count", $binding::GetProperty, $null, $customProps, @())
    for ($i = 1; $i -le $propCount; $i++) {
        $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($i))
        $propName = [System.__ComObject].InvokeMember("Name", $binding::GetProperty, $null, $prop, @())
        $propValue = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, @())
        Write-Host "Property: $propName = $propValue"
    }

    Write-Host "OUT: Check_Custom_Property"
}