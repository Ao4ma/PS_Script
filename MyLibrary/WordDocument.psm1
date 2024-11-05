class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [System.__ComObject]$Document
    [System.__ComObject]$WordApp

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        $this.WordApp = New-Object -ComObject Word.Application
        $this.WordApp.DisplayAlerts = 0  # wdAlertsNone
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.Document = $this.WordApp.Documents.Open($docPath)
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
        $tempFileName = "$($this.DocFileName)_$timestamp"
        $newFilePath = Join-Path -Path $this.DocFilePath -ChildPath $tempFileName
        $this.SaveAs($newFilePath)
        $this.Close()
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
        Rename-Item -Path $newFilePath -NewName (Split-Path $this.DocFileName -Leaf)
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
}

# ドキュメントのパスとファイル名を設定
$DocFileName = "技100-999.docx"
$DocFilePath = "D:\Github\PS_Script"
$ScriptRoot = "D:\Github\PS_Script"

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)

# 必要な操作をここに追加
$wordDoc.SetCustomProperty("CustomPropertyName", "CustomValue")

# ドキュメントを別名で保存してからリネーム
$wordDoc.SetCustomPropertyAndSaveAs("CustomPropertyName", "CustomValue")

# ドキュメントを閉じる
$wordDoc.Close()