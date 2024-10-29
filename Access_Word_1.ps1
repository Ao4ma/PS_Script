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
    }

    # ドキュメントを閉じるメソッド
    [void] Close_Document() {
        if ($null -ne $this.Document) {
            # Ensure $SaveOption is assigned
            $SaveOption = [System.Type]::GetType("Microsoft.Office.Interop.Word.WdSaveOptions")
            $this.Document.Close([ref]$SaveOption::wdDoNotSaveChanges)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
            Remove-Variable -Name Document
        }
        if ($null -ne $this.WordApp) {
            $this.WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
            Remove-Variable -Name WordApp
        }
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
    }

    # COMオブジェクトのメンバーを呼び出すメソッド
    [object] InvokeComObjectMember([object]$ComObject, [string]$MemberName, [string]$Action, [array]$Args = @()) {
        $binding = "System.Reflection.BindingFlags" -as [type]
        try {
            if ($Action -eq "GetProperty") {
                return [System.__ComObject].InvokeMember($MemberName, $binding::GetProperty, $null, $ComObject, $Args)
            } elseif ($Action -eq "SetProperty") {
                [System.__ComObject].InvokeMember($MemberName, $binding::SetProperty, $null, $ComObject, $Args) | Out-Null
                return $null
            } elseif ($Action -eq "InvokeMethod") {
                [System.__ComObject].InvokeMember($MemberName, $binding::InvokeMethod, $null, $ComObject, $Args) | Out-Null
                return $null
            }
        } catch [System.Exception] {
            Write-Host -ForegroundColor Red "Failed to $($Action) $($MemberName): $(($_.Exception).Message)"
            return $null
        }
        return $null
    }

    # ドキュメントプロパティを取得するメソッド
    [object] GetDocumentProperty([object]$Properties, [string]$PropertyName) {
        return $this.InvokeComObjectMember($Properties, "Item", "GetProperty", @($PropertyName))
    }

    # オブジェクトがnullかどうかをチェックするメソッド
    [bool] CheckNull([object]$Object, [string]$ErrorMessage) {
        if ($null -eq $Object) {
            Write-Host -ForegroundColor Red $ErrorMessage
            return $true
        }
        return $false
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

    # ドキュメントを別名で保存するメソッド
    [void] SaveDocumentWithBackup() {
        try {
            $timestamp = Get-Date -Format "yyyyMMddHHmmss"
            $newDocPath = $this.Document.FullName -replace '\.docx$', "_$timestamp.docx"
            $this.Document.SaveAs([ref]$newDocPath)

            # 元のファイルを削除して新しいファイルをリネーム
            $docPath = $this.Document.FullName
            Remove-Item -Path $docPath
            Rename-Item -Path $newDocPath -NewName $docPath
        } catch {
            Write-Host "Failed to save document: $($_)" -ForegroundColor Red
        }
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
            $SaveOption = [System.Type]::GetType("Microsoft.Office.Interop.Word.WdSaveOptions")
            Write-Host "Word library found at $($libraryPath)"
        } else {
            Write-Host "Word library not found at $($libraryPath). Searching the entire system..."
            $libraryPath = Get-ChildItem -Path "C:\" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if ($libraryPath) {
                Add-Type -Path $libraryPath
                $SaveOption = [System.Type]::GetType("Microsoft.Office.Interop.Word.WdSaveOptions")
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
        $customProps = $this.Document.CustomDocumentProperties
        if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot add property.")) {
            Write-Host "OUT: Create_Property"
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
$wordDoc.Create_Property("NewProp2", "NewValue2")
$propValue = $wordDoc.Read_Property("NewProp")
$wordDoc.Update_Property("NewProp", "UpdatedValue")
$wordDoc.Delete_Property("NewProp")
$properties = $wordDoc.Get_Properties("Both")

# ドキュメントを閉じる
$wordDoc.Close_Document()

# Get-Process -Name WINWORD | Stop-Process -Force
# $word = New-Object -ComObject Word.Application
# $doc = $word.Documents.Open($docPath)