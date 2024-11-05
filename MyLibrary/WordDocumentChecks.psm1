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
function Check_Custom_Property {
    param (
        [WordDocument]$wordDoc
    )
    Write-Host "Entering Check_Custom_Property"
    if ($null -eq $wordDoc.Document) {
        Write-Host "Document is null"
    } else {
        $customProps = $wordDoc.Document.CustomDocumentProperties
    }
    if ($null -eq $customProps) {
        Write-Host "customProps is null"
        Write-Host "Exiting Check_Custom_Property"
        return
    }

    $customPropsList = @()
    foreach ($prop in $customProps) {
        $customPropsList += "$($prop.Name): $($prop.Value)"
    }

    $outputFilePath = Join-Path -Path $wordDoc.ScriptRoot -ChildPath "custom_properties.txt"
    Write_ToFile -FilePath $outputFilePath -Content $customPropsList

    Write-Host "Exiting Check_Custom_Property"
}