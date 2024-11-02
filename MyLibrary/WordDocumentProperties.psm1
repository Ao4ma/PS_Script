# MyLibrary/WordDocumentProperties.psm1
# カスタムプロパティを設定するメソッド
[void] SetCustomProperty([string]$PropertyName, [string]$Value) {
    Write-Host "SetCustomProperty: In"
    $customProperties = $this.Document.CustomDocumentProperties
    $binding = "System.Reflection.BindingFlags" -as [type]
    [array]$arrayArgs = $PropertyName, $false, 4, $Value
    try {
        [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        Write-Host "SetCustomProperty: Out"
    } catch [system.exception] {
        try {
            # プロパティが既に存在している場合の処理
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            Write-Host "SetCustomProperty: Out (after delete)"
        } catch {
            Write-Error "Error in SetCustomProperty (inner catch): $_"
            throw $_
        }
    }
}

# カスタムプロパティを読み取るメソッド
[string] Read_Property([string]$propName) {
    Write-Host "IN: Read_Property"
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot read property.")) {
        Write-Host "OUT: Read_Property"
        return $null
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($propName))
    if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
        Write-Host "OUT: Read_Property"
        return $null
    }

    $propValue = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, @())
    Write-Host "OUT: Read_Property"
    return $propValue
}

# カスタムプロパティを更新するメソッド
[void] Update_Property([string]$propName, [string]$propValue) {
    Write-Host "IN: Update_Property"
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot update property.")) {
        Write-Host "OUT: Update_Property"
        return
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($propName))
    if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
        Write-Host "OUT: Update_Property"
        return
    }

    $this.InvokeComObjectMember($prop, "Value", "SetProperty", @($propValue))

    # 一旦別名保存し、クローズして、GCしてからファイルリネーム
    $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
    $this.SaveAs($tempFilePath)
    $this.Close()
    Start-Sleep -Seconds 2  # 少し待機
    Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
    Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)

    Write-Host "OUT: Update_Property"
}

# カスタムプロパティを削除するメソッド
[void] Delete_Property([string]$propName) {
    Write-Host "IN: Delete_Property"
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot delete property.")) {
        Write-Host "OUT: Delete_Property"
        return
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($propName))
    if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
        Write-Host "OUT: Delete_Property"
        return
    }

    $this.InvokeComObjectMember($customProps, "Delete", "InvokeMethod", @($propName))

    # 一旦別名保存し、クローズして、GCしてからファイルリネーム
    $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
    $this.SaveAs($tempFilePath)
    $this.Close()
    Start-Sleep -Seconds 2  # 少し待機
    Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
    Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)

    Write-Host "OUT: Delete_Property"
}

# プロパティを取得するメソッド
[hashtable] Get_Properties([string]$PropertyType) {
    Write-Host "IN: Get_Properties"
    $properties = @{}
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot get properties.")) {
        Write-Host "OUT: Get_Properties"
        return $properties
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $propCount = [System.__ComObject].InvokeMember("Count", $binding::GetProperty, $null, $customProps, @())
    for ($i = 1; $i -le $propCount; $i++) {
        $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($i))
        $propName = [System.__ComObject].InvokeMember("Name", $binding::GetProperty, $null, $prop, @())
        $propValue = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, @())
        $properties[$propName] = $propValue
    }

    Write-Host "OUT: Get_Properties"
    return $properties
}