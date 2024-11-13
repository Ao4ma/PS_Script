using module .\WordDocumentUtilities.psm1
<#
function Check_PC_Env {
    param (
        [WordDocument]$wordDoc
    )
    Write-Host "IN: Check_PC_Env"
    $envInfo = @{
        "PCName" = $env:COMPUTERNAME
        "PowerShellHome" = $env:PSHOME
        "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
        "MACAddress" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).MacAddress
    }
    $envInfo

    $filePath = Join-Path -Path $wordDoc.ScriptRoot -ChildPath "$($env:COMPUTERNAME)_env_info.txt"
    Write_ToFile -FilePath $filePath -Content ($envInfo.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" })

    Write-Host "OUT: Check_PC_Env"
}

function Check_Word_Library {
    Write-Host "IN: Check_Word_Library"
    $libraryPath = Get-ChildItem -Path "C:\" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($libraryPath) {
        Add-Type -Path $libraryPath
        Write-Host "Word library found at $($libraryPath)"
    } else {
        Write-Host -ForegroundColor Red "Word library not found on this system."
        throw "Word library not found. Please install the required library."
    }
    Write-Host "OUT: Check_Word_Library"
}
#>
function checkCustomProperty {
    param (
        [WordDocument]$wordDoc
    )
    Write-Host "Entering checkCustomProperty"
    if ($null -eq $wordDoc.Document) {
        Write-Host "Document is null"
    } else {
        $customProps = $wordDoc.Document.CustomDocumentProperties
    }
    if ($null -eq $customProps) {
        Write-Host "customProps is null"
        Write-Host "Exiting checkCustomProperty"
        return
    }

    $customPropsList = @()
    $bindingFlags = [System.Reflection.BindingFlags]::GetProperty -bor [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::Public

    foreach ($prop in $customProps) {
        $propName = $prop.Name
        $propValue = $prop.GetType().InvokeMember("Value", $bindingFlags, $null, $prop, $null)
        $customPropsList += "$propName: $propValue"
    }

    $outputFilePath = Join-Path -Path $wordDoc.ScriptRoot -ChildPath "custom_properties.txt"
    Write_ToFile -FilePath $outputFilePath -Content $customPropsList

    Write-Host "Exiting checkCustomProperty"
}