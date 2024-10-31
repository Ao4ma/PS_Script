class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [object]$WordApp
    [object]$Document

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        $this.Open_Document()
    }

    # ドキュメントを開くメソッド
    [void] Open_Document() {
        $this.WordApp = New-Object -ComObject Word.Application
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.Document = $this.WordApp.Documents.Open($docPath)
        if ($null -eq $this.Document) {
            throw "Failed to open document: $docPath"
        }
    }

    # ドキュメントを閉じるメソッド
    [void] Close_Document() {
        Write-Host "IN: Close_Document"
        if ($null -ne $this.Document) {
            $this.Document.Close()
            $this.Document = $null
        }
        if ($null -ne $this.WordApp) {
            $this.WordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
            $this.WordApp = $null
        }
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
        Write-Host "OUT: Close_Document"
    }

    # ドキュメントを別名で保存するメソッド
    [void] SaveDocumentWithBackup() {
        try {
            if ($null -eq $this.Document.FullName) {
                throw "Document path is null. Cannot save document."
            }

            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newDocPath = $this.Document.FullName -replace '\.docx$', "_$timestamp.docx"
            $this.Document.SaveAs([ref]$newDocPath)

            # デバッグ用メッセージ
            Write-Host "Document saved as: $newDocPath"

            # 保存直後にファイルの存在を確認
            $retryCount = 5
            $retryInterval = 2 # seconds
            for ($i = 0; $i -lt $retryCount; $i++) {
                if (Test-Path -Path $newDocPath) {
                    Write-Host "New document path exists: $newDocPath"
                    break
                } else {
                    Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist immediately after save. Retrying..."
                    Start-Sleep -Seconds $retryInterval
                }
            }

            if (-not (Test-Path -Path $newDocPath)) {
                throw "New document path '$newDocPath' does not exist after ${retryCount} retries. Save operation failed."
            }

            # Close the document and release the COM objects
            $originalDocPath = $this.Document.FullName
            $this.Close_Document()

            # リトライ設定
            $retryCount = 5
            $retryInterval = 2 # seconds

            # ファイルの存在を確認し、リトライ
            for ($i = 0; $i -lt $retryCount; $i++) {
                if (Test-Path -Path $newDocPath) {
                    Write-Host "New document path exists: $newDocPath"
                    Remove-Item -Path $originalDocPath
                    Start-Sleep -Seconds 1 # 少し待機してからリネーム
                    if (Test-Path -Path $newDocPath) {
                        Rename-Item -Path $newDocPath -NewName $originalDocPath
                        Write-Host "Document renamed successfully."
                        return
                    } else {
                        Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist after waiting. Retrying..."
                        Start-Sleep -Seconds $retryInterval
                    }
                } else {
                    Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist. Retrying..."
                    Start-Sleep -Seconds $retryInterval
                }
            }

            throw "New document path '$newDocPath' does not exist after ${retryCount} retries. Rename operation failed."
        } catch {
            Write-Host "Failed to save document: $($_)" -ForegroundColor Red
        }
    }

    # ファイルに書き込むメソッド
    [void] WriteToFile([string]$filePath, [string]$content) {
        Set-Content -Path $filePath -Value $content
    }

    # Nullチェックメソッド
    [bool] CheckNull([object]$obj, [string]$message) {
        if ($null -eq $obj) {
            Write-Host $message -ForegroundColor Red
            return $true
        }
        return $false
    }

    # カスタムプロパティを取得するメソッド
    [object] GetDocumentProperty([object]$properties, [string]$propertyName) {
        foreach ($property in $properties) {
            if ($property.Name -eq $propertyName) {
                return $property
            }
        }
        return $null
    }

    # COMオブジェクトのメンバーを呼び出すメソッド
    [object] InvokeComObjectMember([object]$comObject, [string]$memberName, [string]$memberType, [object[]]$args) {
        if ($null -eq $comObject) {
            throw "COM object is null. Cannot invoke member."
        }
        $bindingFlags = [System.Reflection.BindingFlags]::InvokeMethod
        if ($memberType -eq "GetProperty") {
            $bindingFlags = [System.Reflection.BindingFlags]::GetProperty
        } elseif ($memberType -eq "SetProperty") {
            $bindingFlags = [System.Reflection.BindingFlags]::SetProperty
        }
        return $comObject.GetType().InvokeMember($memberName, $bindingFlags, $null, $comObject, $args)
    }

    # PC環境をチェックするメソッド
    [void] Check_PC_Env() {
        Write-Host "IN: Check_PC_Env"
        $envInfo = @{
            "PCName" = $env:COMPUTERNAME
            "PowerShellHome" = $env:PSHOME
            "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
            "MACAddress" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).MacAddress
            "DocFilePath" = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            "ScriptLibraryPath" = $this.ScriptRoot
        }
        $envInfo

        # ファイルに出力
        $filePath = Join-Path -Path $this.ScriptRoot -ChildPath "$($env:COMPUTERNAME)_env_info.txt"
        $this.WriteToFile($filePath, ($envInfo.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }))

        Write-Host "OUT: Check_PC_Env"
    }

    # Wordライブラリをチェックするメソッド
    [void] Check_Word_Library() {
        Write-Host "IN: Check_Word_Library"
        $libraryPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
        if (Test-Path $libraryPath) {
            Add-Type -Path $libraryPath
            Write-Host "Word library found at $($libraryPath)"
        } else {
            Write-Host "Word library not found at $($libraryPath). Searching the entire system..."
            $libraryPath = Get-ChildItem -Path "C:\" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if ($libraryPath) {
                Add-Type -Path $libraryPath
                Write-Host "Word library found at $($libraryPath)"
            } else {
                Write-Host -ForegroundColor Red "Word library not found on this system."
                throw "Word library not found. Please install the required library."
            }
        }
        Write-Host "OUT: Check_Word_Library"
    }

    # カスタムプロパティをチェックするメソッド
    [void] Check_Custom_Property() {
        Write-Host "Entering Check_Custom_Property"
        $customProps = $this.Document.CustomDocumentProperties
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
        $this.WriteToFile($outputFilePath, $customPropsList)

        Write-Host "Exiting Check_Custom_Property"
    }

    # カスタムプロパティを作成するメソッド
    [void] Create_Property([string]$propName, [string]$propValue) {
        Write-Host "IN: Create_Property"

        # ドキュメントが開かれているか確認
        if ($null -eq $this.Document) {
            throw "Document is not open. Cannot create property."
        }

        $customProps = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot add property.")) {
            return
        }

        [array]$arrayArgs = $propName, $false, 4, $propValue
        $this.InvokeComObjectMember($customProps, "Add", "InvokeMethod", $arrayArgs)
        $this.SaveDocumentWithBackup()
        Write-Host "OUT: Create_Property"
    }

    # カスタムプロパティを読み取るメソッド
    [string] Read_Property([string]$propName) {
        Write-Host "IN: Read_Property"
        $customProps = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot read property.")) {
            Write-Host "OUT: Read_Property"
            return $null
        }

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Read_Property"
            return $null
        }

        $propValue = $this.InvokeComObjectMember($prop, "Value", "GetProperty", @())
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

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Update_Property"
            return
        }

        $this.InvokeComObjectMember($prop, "Value", "SetProperty", @($propValue))
        $this.SaveDocumentWithBackup()
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

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Delete_Property"
            return
        }

        $this.InvokeComObjectMember($prop, "Delete", "InvokeMethod", @())
        $this.SaveDocumentWithBackup()
        Write-Host "OUT: Delete_Property"
    }

    # プロパティを取得するメソッド
    [hashtable] Get_Properties([string]$PropertyType) {
        Write-Host "IN: Get_Properties"
        $this.BuiltinPropertiesGroup = @(
            "Title", "Subject", "Author", "Keywords", "Comments", "Template", "Last Author", 
            "Revision Number", "Application Name", "Last Print Date", "Creation Date", 
            "Last Save Time", "Total Editing Time", "Number of Pages", "Number of Words", 
            "Number of Characters", "Security", "Category", "Format", "Manager", "Company", 
            "Number of Bytes", "Number of Lines", "Number of Paragraphs", "Number of Slides", 
            "Number of Notes", "Number of Hidden Slides", "Number of Multimedia Clips", 
            "Hyperlink Base", "Number of Characters (with spaces)", "Content Type", 
            "Content Status", "Language", "Document Version"
        )

        $this.CustomPropertiesGroup = @("batter", "yamada") # これは例

        $objHash = @{}
        $foundProperties = @()

        $this.GetPropertiesByType($this.Document.BuiltInDocumentProperties, $this.BuiltinPropertiesGroup, $objHash, $foundProperties, $PropertyType, "Builtin")
        $this.GetPropertiesByType($this.Document.CustomDocumentProperties, $this.CustomPropertiesGroup, $objHash, $foundProperties, $PropertyType, "Custom")

        Write-Host "OUT: Get_Properties"
        return $objHash
    }

    # プロパティの種類に応じてプロパティを取得するメソッド
    [void] GetPropertiesByType([object]$Properties, [array]$PropertiesGroup, [hashtable]$objHash, [array]$foundProperties, [string]$PropertyType, [string]$Type) {
        if ($PropertyType -eq $Type -or $PropertyType -eq "Both") {
            foreach ($p in $PropertiesGroup) {
                $value = $this.GetDocumentProperty($Properties, $p)
                if ($null -ne $value) {
                    $objHash[$p] = $value
                    $foundProperties += $p
                }
            }
        }
    }

    # カスタムプロパティを設定して別名で保存するメソッド
    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value, [string]$NewFilePath) {
        Write-Host "IN: SetCustomPropertyAndSaveAs"
        $this.Create_Property($PropertyName, $Value)
        $this.Document.SaveAs([ref]$NewFilePath)
        $this.Close_Document()
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
        Rename-Item -Path $NewFilePath -NewName (Split-Path $this.DocFileName -Leaf)
        Write-Host "OUT: SetCustomPropertyAndSaveAs"
    }

    # Wordプロセスを閉じるメソッド
    [void] Close_Word_Processes() {
        Write-Host "IN: Close_Word_Processes"
        $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
        if ($existingWordProcesses) {
            foreach ($process in $existingWordProcesses) {
                Stop-Process -Id $process.Id -Force
            }
        }
        Write-Host "OUT: Close_Word_Processes"
    }

    # Wordが閉じられていることを確認するメソッド
    [void] Ensure_Word_Closed() {
        Write-Host "IN: Ensure_Word_Closed"
        $newWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
        if ($newWordProcesses) {
            foreach ($process in $newWordProcesses) {
                Stop-Process -Id $process.Id -Force
            }
        }
        Write-Host "OUT: Ensure_Word_Closed"
    }
}

# デバッグ用設定
$DocFileName = "技100-999.docx"
$ScriptRoot1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$ScriptRoot2 = "D:\Github\PS_Script"

# デバッグ環境に応じてパスを切り替える
if (Test-Path $ScriptRoot2) {
    $ScriptRoot = $ScriptRoot2
} else {
    $ScriptRoot = $ScriptRoot1
}
$DocFilePath = $ScriptRoot

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)

# メソッドの呼び出し例
$wordDoc.Check_PC_Env()
$wordDoc.Check_Word_Library()
$wordDoc.Check_Custom_Property()
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty2", "Value2", "D:\Github\PS_Script\sample_temp.docx")

# ドキュメントを閉じる
$wordDoc.Closeclass WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [object]$WordApp
    [object]$Document

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
        $this.Open_Document()
    }

    # ドキュメントを開くメソッド
    [void] Open_Document() {
        $this.WordApp = New-Object -ComObject Word.Application
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $this.Document = $this.WordApp.Documents.Open($docPath)
        if ($null -eq $this.Document) {
            throw "Failed to open document: $docPath"
        }
    }

    # ドキュメントを閉じるメソッド
    [void] Close_Document() {
        Write-Host "IN: Close_Document"
        if ($null -ne $this.Document) {
            $this.Document.Close()
            $this.Document = $null
        }
        if ($null -ne $this.WordApp) {
            $this.WordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
            $this.WordApp = $null
        }
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
        Write-Host "OUT: Close_Document"
    }

    # ドキュメントを別名で保存するメソッド
    [void] SaveDocumentWithBackup() {
        try {
            if ($null -eq $this.Document.FullName) {
                throw "Document path is null. Cannot save document."
            }

            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newDocPath = $this.Document.FullName -replace '\.docx$', "_$timestamp.docx"
            $this.Document.SaveAs([ref]$newDocPath)

            # デバッグ用メッセージ
            Write-Host "Document saved as: $newDocPath"

            # 保存直後にファイルの存在を確認
            $retryCount = 5
            $retryInterval = 2 # seconds
            for ($i = 0; $i -lt $retryCount; $i++) {
                if (Test-Path -Path $newDocPath) {
                    Write-Host "New document path exists: $newDocPath"
                    break
                } else {
                    Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist immediately after save. Retrying..."
                    Start-Sleep -Seconds $retryInterval
                }
            }

            if (-not (Test-Path -Path $newDocPath)) {
                throw "New document path '$newDocPath' does not exist after ${retryCount} retries. Save operation failed."
            }

            # Close the document and release the COM objects
            $originalDocPath = $this.Document.FullName
            $this.Close_Document()

            # リトライ設定
            $retryCount = 5
            $retryInterval = 2 # seconds

            # ファイルの存在を確認し、リトライ
            for ($i = 0; $i -lt $retryCount; $i++) {
                if (Test-Path -Path $newDocPath) {
                    Write-Host "New document path exists: $newDocPath"
                    Remove-Item -Path $originalDocPath
                    Start-Sleep -Seconds 1 # 少し待機してからリネーム
                    if (Test-Path -Path $newDocPath) {
                        Rename-Item -Path $newDocPath -NewName $originalDocPath
                        Write-Host "Document renamed successfully."
                        return
                    } else {
                        Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist after waiting. Retrying..."
                        Start-Sleep -Seconds $retryInterval
                    }
                } else {
                    Write-Host "Retry $($i + 1)/${retryCount}: New document path '$newDocPath' does not exist. Retrying..."
                    Start-Sleep -Seconds $retryInterval
                }
            }

            throw "New document path '$newDocPath' does not exist after ${retryCount} retries. Rename operation failed."
        } catch {
            Write-Host "Failed to save document: $($_)" -ForegroundColor Red
        }
    }

    # ファイルに書き込むメソッド
    [void] WriteToFile([string]$filePath, [string]$content) {
        Set-Content -Path $filePath -Value $content
    }

    # Nullチェックメソッド
    [bool] CheckNull([object]$obj, [string]$message) {
        if ($null -eq $obj) {
            Write-Host $message -ForegroundColor Red
            return $true
        }
        return $false
    }

    # カスタムプロパティを取得するメソッド
    [object] GetDocumentProperty([object]$properties, [string]$propertyName) {
        foreach ($property in $properties) {
            if ($property.Name -eq $propertyName) {
                return $property
            }
        }
        return $null
    }

    # COMオブジェクトのメンバーを呼び出すメソッド
    [object] InvokeComObjectMember([object]$comObject, [string]$memberName, [string]$memberType, [object[]]$args) {
        if ($null -eq $comObject) {
            throw "COM object is null. Cannot invoke member."
        }
        $bindingFlags = [System.Reflection.BindingFlags]::InvokeMethod
        if ($memberType -eq "GetProperty") {
            $bindingFlags = [System.Reflection.BindingFlags]::GetProperty
        } elseif ($memberType -eq "SetProperty") {
            $bindingFlags = [System.Reflection.BindingFlags]::SetProperty
        }
        return $comObject.GetType().InvokeMember($memberName, $bindingFlags, $null, $comObject, $args)
    }

    # PC環境をチェックするメソッド
    [void] Check_PC_Env() {
        Write-Host "IN: Check_PC_Env"
        $envInfo = @{
            "PCName" = $env:COMPUTERNAME
            "PowerShellHome" = $env:PSHOME
            "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
            "MACAddress" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).MacAddress
            "DocFilePath" = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            "ScriptLibraryPath" = $this.ScriptRoot
        }
        $envInfo

        # ファイルに出力
        $filePath = Join-Path -Path $this.ScriptRoot -ChildPath "$($env:COMPUTERNAME)_env_info.txt"
        $this.WriteToFile($filePath, ($envInfo.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }))

        Write-Host "OUT: Check_PC_Env"
    }

    # Wordライブラリをチェックするメソッド
    [void] Check_Word_Library() {
        Write-Host "IN: Check_Word_Library"
        $libraryPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
        if (Test-Path $libraryPath) {
            Add-Type -Path $libraryPath
            Write-Host "Word library found at $($libraryPath)"
        } else {
            Write-Host "Word library not found at $($libraryPath). Searching the entire system..."
            $libraryPath = Get-ChildItem -Path "C:\" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if ($libraryPath) {
                Add-Type -Path $libraryPath
                Write-Host "Word library found at $($libraryPath)"
            } else {
                Write-Host -ForegroundColor Red "Word library not found on this system."
                throw "Word library not found. Please install the required library."
            }
        }
        Write-Host "OUT: Check_Word_Library"
    }

    # カスタムプロパティをチェックするメソッド
    [void] Check_Custom_Property() {
        Write-Host "Entering Check_Custom_Property"
        $customProps = $this.Document.CustomDocumentProperties
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
        $this.WriteToFile($outputFilePath, $customPropsList)

        Write-Host "Exiting Check_Custom_Property"
    }

    # カスタムプロパティを作成するメソッド
    [void] Create_Property([string]$propName, [string]$propValue) {
        Write-Host "IN: Create_Property"

        # ドキュメントが開かれているか確認
        if ($null -eq $this.Document) {
            throw "Document is not open. Cannot create property."
        }

        $customProps = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot add property.")) {
            return
        }

        [array]$arrayArgs = $propName, $false, 4, $propValue
        $this.InvokeComObjectMember($customProps, "Add", "InvokeMethod", $arrayArgs)
        $this.SaveDocumentWithBackup()
        Write-Host "OUT: Create_Property"
    }

    # カスタムプロパティを読み取るメソッド
    [string] Read_Property([string]$propName) {
        Write-Host "IN: Read_Property"
        $customProps = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot read property.")) {
            Write-Host "OUT: Read_Property"
            return $null
        }

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Read_Property"
            return $null
        }

        $propValue = $this.InvokeComObjectMember($prop, "Value", "GetProperty", @())
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

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Update_Property"
            return
        }

        $this.InvokeComObjectMember($prop, "Value", "SetProperty", @($propValue))
        $this.SaveDocumentWithBackup()
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

        $prop = $this.GetDocumentProperty($customProps, $propName)
        if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
            Write-Host "OUT: Delete_Property"
            return
        }

        $this.InvokeComObjectMember($prop, "Delete", "InvokeMethod", @())
        $this.SaveDocumentWithBackup()
        Write-Host "OUT: Delete_Property"
    }

    # プロパティを取得するメソッド
    [hashtable] Get_Properties([string]$PropertyType) {
        Write-Host "IN: Get_Properties"
        $this.BuiltinPropertiesGroup = @(
            "Title", "Subject", "Author", "Keywords", "Comments", "Template", "Last Author", 
            "Revision Number", "Application Name", "Last Print Date", "Creation Date", 
            "Last Save Time", "Total Editing Time", "Number of Pages", "Number of Words", 
            "Number of Characters", "Security", "Category", "Format", "Manager", "Company", 
            "Number of Bytes", "Number of Lines", "Number of Paragraphs", "Number of Slides", 
            "Number of Notes", "Number of Hidden Slides", "Number of Multimedia Clips", 
            "Hyperlink Base", "Number of Characters (with spaces)", "Content Type", 
            "Content Status", "Language", "Document Version"
        )

        $this.CustomPropertiesGroup = @("batter", "yamada") # これは例

        $objHash = @{}
        $foundProperties = @()

        $this.GetPropertiesByType($this.Document.BuiltInDocumentProperties, $this.BuiltinPropertiesGroup, $objHash, $foundProperties, $PropertyType, "Builtin")
        $this.GetPropertiesByType($this.Document.CustomDocumentProperties, $this.CustomPropertiesGroup, $objHash, $foundProperties, $PropertyType, "Custom")

        Write-Host "OUT: Get_Properties"
        return $objHash
    }

    # プロパティの種類に応じてプロパティを取得するメソッド
    [void] GetPropertiesByType([object]$Properties, [array]$PropertiesGroup, [hashtable]$objHash, [array]$foundProperties, [string]$PropertyType, [string]$Type) {
        if ($PropertyType -eq $Type -or $PropertyType -eq "Both") {
            foreach ($p in $PropertiesGroup) {
                $value = $this.GetDocumentProperty($Properties, $p)
                if ($null -ne $value) {
                    $objHash[$p] = $value
                    $foundProperties += $p
                }
            }
        }
    }

    # カスタムプロパティを設定して別名で保存するメソッド
    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value, [string]$NewFilePath) {
        Write-Host "IN: SetCustomPropertyAndSaveAs"
        $this.Create_Property($PropertyName, $Value)
        $this.Document.SaveAs([ref]$NewFilePath)
        $this.Close_Document()
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
        Rename-Item -Path $NewFilePath -NewName (Split-Path $this.DocFileName -Leaf)
        Write-Host "OUT: SetCustomPropertyAndSaveAs"
    }

    # Wordプロセスを閉じるメソッド
    [void] Close_Word_Processes() {
        Write-Host "IN: Close_Word_Processes"
        $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
        if ($existingWordProcesses) {
            foreach ($process in $existingWordProcesses) {
                Stop-Process -Id $process.Id -Force
            }
        }
        Write-Host "OUT: Close_Word_Processes"
    }

    # Wordが閉じられていることを確認するメソッド
    [void] Ensure_Word_Closed() {
        Write-Host "IN: Ensure_Word_Closed"
        $newWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
        if ($newWordProcesses) {
            foreach ($process in $newWordProcesses) {
                Stop-Process -Id $process.Id -Force
            }
        }
        Write-Host "OUT: Ensure_Word_Closed"
    }
}

# デバッグ用設定
$DocFileName = "技100-999.docx"
$ScriptRoot1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$ScriptRoot2 = "D:\Github\PS_Script"

# デバッグ環境に応じてパスを切り替える
if (Test-Path $ScriptRoot2) {
    $ScriptRoot = $ScriptRoot2
} else {
    $ScriptRoot = $ScriptRoot1
}
$DocFilePath = $ScriptRoot

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)

# メソッドの呼び出し例
$wordDoc.Check_PC_Env()
$wordDoc.Check_Word_Library()
$wordDoc.Check_Custom_Property()
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty2", "Value2", "D:\Github\PS_Script\sample_temp.docx")

$wordDoc.Check_Custom_Property()
$propValue = $wordDoc.Read_Property("NewProp")
$wordDoc.Update_Property("NewProp", "UpdatedValue")
$wordDoc.Delete_Property("NewProp")

# ドキュメントを閉じる
$wordDoc.Close_Document()

# Get-Process -Name WINWORD | Stop-Process -Force