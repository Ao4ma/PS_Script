<#
function Set_CustomProperty {
    param (
        [WordDocument]$wordDoc,
        [string]$PropertyName,
        [string]$Value
    )
    Write-Host "Set_CustomProperty: In"
    $customProperties = $wordDoc.Document.CustomDocumentProperties
    $binding = "System.Reflection.BindingFlags" -as [type]
    [array]$arrayArgs = $PropertyName, $false, 4, $Value
    try {
        [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        Write-Host "Set_CustomProperty: Out"
    } catch [system.exception] {
        try {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            Write-Host "Set_CustomProperty: Out (after delete)"
        } catch {
            Write-Error "Error in Set_CustomProperty (inner catch): $_"
            throw $_
        }
    }
}
#>
function Read_Property {
    param (
        [WordDocument]$wordDoc,
        [string]$PropertyName
    )
    Write-Host "IN: Read_Property"
    $customProps = $wordDoc.Document.CustomDocumentProperties
    if ($null -eq $customProps) {
        Write-Host "CustomDocumentProperties is null. Cannot read property."
        Write-Host "OUT: Read_Property"
        return $null
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($PropertyName))
    if ($null -eq $prop) {
        Write-Host "Property '$($PropertyName)' not found."
        Write-Host "OUT: Read_Property"
        return $null
    }

    $propValue = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, @())
    Write-Host "OUT: Read_Property"
    return $propValue
}

function Update_Property {
    param (
        [WordDocument]$wordDoc,
        [string]$PropertyName,
        [string]$PropertyValue
    )
    Write-Host "IN: Update_Property"
    $customProps = $wordDoc.Document.CustomDocumentProperties
    if ($null -eq $customProps) {
        Write-Host "CustomDocumentProperties is null. Cannot update property."
        Write-Host "OUT: Update_Property"
        return
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($PropertyName))
    if ($null -eq $prop) {
        Write-Host "Property '$($PropertyName)' not found."
        Write-Host "OUT: Update_Property"
        return
    }

    [System.__ComObject].InvokeMember("Value", $binding::SetProperty, $null, $prop, @($PropertyValue))
    Write-Host "OUT: Update_Property"
}

function Delete_Property {
    param (
        [WordDocument]$wordDoc,
        [string]$PropertyName
    )
    Write-Host "IN: Delete_Property"
    $customProps = $wordDoc.Document.CustomDocumentProperties
    if ($null -eq $customProps) {
        Write-Host "CustomDocumentProperties is null. Cannot delete property."
        Write-Host "OUT: Delete_Property"
        return
    }

    $binding = "System.Reflection.BindingFlags" -as [type]
    $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($PropertyName))
    if ($null -eq $prop) {
        Write-Host "Property '$($PropertyName)' not found."
        Write-Host "OUT: Delete_Property"
        return
    }

    [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $prop, $null)
    Write-Host "OUT: Delete_Property"
}