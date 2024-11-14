using module .\WordDocumentUtilities.psm1

class WordDocument {
    [string]$docFilePath
    [string]$scriptRoot
    [System.__ComObject]$document
    [System.__ComObject]$wordApp

    WordDocument([string]$docFilePath, [string]$scriptRoot) {
        # $this.DocFileName
        # $this.docFilePath = $docFilePath
        $this.scriptRoot = $scriptRoot
        $this.wordApp = New-Object -ComObject Word.Application
        $this.wordApp.DisplayAlerts = 0  # wdAlertsNone
        $this.docFilePath = $docFilePath
        if (-not (Test-Path $docFilePath)) {
            throw "ドキュメントが見つかりません: $docFilePath"
        }
        $this.Document = $this.WordApp.Documents.Open($docFilePath)
    }

    # ドキュメントを閉じるメソッド
    [void] Close() {
        if ($null -ne $this.Document) {
            $this.Document.Close([ref]0)  # 0はwdDoNotSaveChangesに相当
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
            # Remove-Variable -Name Document
            $this.Document = $null
        }
        if ($null -ne $this.WordApp) {
            $this.WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
            # Remove-Variable -Name WordApp
            $this.WordApp = $null
        }
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
    }

    # ドキュメントを別名で保存するメソッド
    [void] SaveAs([string]$newFilePath) {
        $this.Document.SaveAs([ref]$newFilePath)
    }

    # カスタムプロパティを設定するメソッド
    [void] SetCustomProperty([string]$PropertyName, [string]$Value) {
        $customProperties = $this.Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName, $false, 4, $Value
        try {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        } catch [system.exception] {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }
    }

    # カスタムプロパティを設定して別名で保存するメソッド
    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value) {
        Write-Host "SetCustomPropertyAndSaveAs: In"
        $this.SetCustomProperty($PropertyName, $Value)
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($this.DocFilePath)
        $extension = [System.IO.Path]::GetExtension($this.DocFilePath)
        $tempFileName = "$($baseName)_$($timestamp)$($extension)"
        $newFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($this.DocFilePath), $tempFileName)
        Write-Host $newFilePath
        
        $this.SaveAs($newFilePath)
        $this.Close()
        Start-Sleep -Seconds 1  # 少し待機
        Remove-Item -Path $this.DocFilePath -Force
        # Start-Sleep -Seconds 1  # 少し待機 
        Rename-Item -Path $newFilePath -NewName (Split-Path $this.DocFilePath -Leaf)
        Write-Host "SetCustomPropertyAndSaveAs: Out"
    }

    # PC環境をチェックするメソッド
    [void] Check_PC_Env() {
        $envInfo = @{
            "PCName" = $env:COMPUTERNAME
            "PowerShellHome" = $env:PSHOME
            "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
            "MACAddress" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).MacAddress
            "DocFilePath" = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            "ScriptLibraryPath" = $this.ScriptRoot
        }
        Write-Host "PC Environment Info: $envInfo"
    }

    # Wordライブラリをチェックするメソッド
    [void] Check_Word_Library() {
        $libraryPath = Join-Path -Path $this.ScriptRoot -ChildPath "WordLibrary.dll"
        if (Test-Path $libraryPath) {
            Add-Type -Path $libraryPath
            Write-Host "Word library found at $($libraryPath)"
        } else {
            Write-Host "Word library not found at $($libraryPath). Searching the entire system..."
        }
    }

    # カスタムプロパティをチェックするメソッド
    [void] checkCustomProperty2() {
        Write-Host "Entering Check_Custom_Property"
        if ($null -eq $this.Document) {
            Write-Host "Document is null"
            Write-Host "Exiting Check_Custom_Property"
            return
        } else {
            $customProps = $this.Document.CustomDocumentProperties
        }
        if ($this.CheckNull($customProps, "customProps is null")) {
            Write-Host "Exiting Check_Custom_Property"
            return
        }

        $customPropsList = @()
        foreach ($prop in $customProps) {
            $propName = $this.InvokeComObjectMember($prop, "Name", "GetProperty", @())
            if ($null -ne $propName) {
                $customPropsList += $propName
            }
        }

        # ファイルに出力
        $outputFilePath = Join-Path -Path $this.ScriptRoot -ChildPath "custom_properties.txt"
        this.WriteToFile($outputFilePath, $customPropsList)

        Write-Host "Exiting Check_Custom_Property"
    }



    # Nullチェックメソッド
    [bool] CheckNull([object]$obj, [string]$message) {
        if ($null -eq $obj) {
            Write-Host $message -ForegroundColor Red
            return $true
        }
        return $false
    }

    # カスタムプロパティを読み取るメソッド
    [object] Read_Property2([string]$PropertyName) {
        Write-Host "IN: Read_Property"
        $customProperties = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProperties, "CustomDocumentProperties is null. Cannot read property.")) {
            Write-Host "OUT: Read_Property"
            return $null
        }
    
        $binding = "System.Reflection.BindingFlags" -as [type]
        try {
            $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, @($PropertyName))
            $propValue = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $prop, $null)
            Write-Host "Read Property Value: $propValue"
            Write-Host "OUT: Read_Property"
            return $propValue
        } catch {
            Write-Error "Error in Read_Property: $_"
            Write-Host "OUT: Read_Property"
            return $null
        }
    }

    # ファイルに内容を書き込むメソッド
    [void] WriteToFile([string]$FilePath, [array]$Content) {
        if ($Content.Count -eq 0) {
            Write-Host "No content found. Deleting previous output file if it exists."
            if (Test-Path $FilePath) {
                Remove-Item $FilePath
            }
        } else {
            $Content | Out-File -FilePath $FilePath
        }
    }

}

<# ドキュメントのパスとファイル名を設定
$DocFileName = "技100-999.docx"
$DocFilePath = "D:\Github\PS_Script"
$ScriptRoot = "D:\Github\PS_Script"

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)

# 必要な操作をここに追加
$wordDoc.SetCustomProperty("CustomPropertyName", "CustomValue")

# ドキュメントを別名で保存してからリネーム
$wordDoc.SetCustomPropertyAndSaveAs("CustomPropertyName", "CustomValue")

# ドキュメントを閉じる
$wordDoc.Close()
#>