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

    # ドキュメントを別名で保存するメソッド
    [void] SaveAs([string]$NewFilePath) {
        $this.Document.SaveAs([ref]$NewFilePath)
    }

    # ドキュメントを閉じるメソッド
    [void] Close() {
        $this.Document.Close([ref]$false)  # false を指定して変更を保存しない
        $this.WordApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # カスタムプロパティを設定して別名で保存するメソッド
    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value, [string]$NewFilePath) {
        Write-Host "SetCustomPropertyAndSaveAs: In"
        $this.SetCustomProperty($PropertyName, $Value)
        $this.SaveAs($NewFilePath)
        $this.Close()
        Start-Sleep -Seconds 2  # 少し待機
        Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
        Rename-Item -Path $NewFilePath -NewName (Split-Path $this.DocFileName -Leaf)
        Write-Host "SetCustomPropertyAndSaveAs: Out"
    }

    # Nullチェックメソッド
    [bool] CheckNull([object]$obj, [string]$message) {
        if ($null -eq $obj) {
            Write-Host $message -ForegroundColor Red
            return $true
        }
        return $false
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

    # InvokeComObjectMember メソッドを追加
    [object] InvokeComObjectMember([object]$comObject, [string]$memberName, [string]$bindingFlags, [array]$args) {
        $binding = "System.Reflection.BindingFlags" -as [type]
        return [System.__ComObject].InvokeMember($memberName, $binding::$bindingFlags, $null, $comObject, $args)
    }

    # カスタムプロパティをチェックするメソッド
    [void] Check_Custom_Property() {
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
        $this.WriteToFile($outputFilePath, $customPropsList)

        Write-Host "Exiting Check_Custom_Property"
    }

    # カスタムプロパティを読み取るメソッド
    [object] Read_Property([string]$PropertyName) {
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
[void] Delete_Property([string]$PropertyName) {
    Write-Host "IN: Delete_Property"
    $customProps = $this.Document.CustomDocumentProperties
    if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot delete property.")) {
        Write-Host "OUT: Delete_Property"
        return
    }
    $binding = "System.Reflection.BindingFlags" -as [type]
    try {
        $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($PropertyName))
        [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
    } catch {
        Write-Error "Error in Delete_Property: $_"
        Write-Host "OUT: Delete_Property"
    }
    
    # 一旦別名保存し、クローズして、GCしてからファイルリネーム
    $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
    $this.SaveAs($tempFilePath)
    $this.Close()
    Start-Sleep -Seconds 2  # 少し待機
    Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
    Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)

    # 新しいインスタンスを作成してから操作を続行
    $this.WordApp = New-Object -ComObject Word.Application
    $this.WordApp.DisplayAlerts = 0  # wdAlertsNone
    $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
    $this.Document = $this.WordApp.Documents.Open($docPath)
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

    # サイン欄に名前と日付を配置するメソッド
    [void] FillSignatures() {
        Write-Host "IN: FillSignatures"

        # カスタムプロパティから担当者、承認者、照査者の名前と日付を取得
        $担当者 = $this.Read_Property("担当者")
        $担当者日付 = $this.Read_Property("担当者日付")
        $承認者 = $this.Read_Property("承認者")
        $承認者日付 = $this.Read_Property("承認者日付")
        $照査者 = $this.Read_Property("照査者")
        $照査者日付 = $this.Read_Property("照査者日付")

        # フォントサイズを計算する関数
        function CalculateFontSize($text, $cellWidth) {
            $averageCharWidth = 0.6 # 平均的な文字の幅の割合（文字数に基づく調整）
            $textLength = $text.Length
            $fontSize = [math]::Floor($cellWidth / ($averageCharWidth * $textLength))
            return $fontSize - 1 # 少し余裕を持たせるために1ポイント減らす
        }

        # 表のサイン欄に名前と日付を配置
        $table = $this.Document.Tables.Item(1) # 1番目の表を取得
        $cell1 = $table.Cell(2, 1) # 2行1列目
        $cell2 = $table.Cell(2, 2) # 2行2列目
        $cell3 = $table.Cell(2, 3) # 2行3列目

        # サイン欄の横幅を取得
        $cellWidth1 = $cell1.Width
        $cellWidth2 = $cell2.Width
        $cellWidth3 = $cell3.Width

        # フォントサイズを計算
        $fontSize1 = CalculateFontSize($担当者, $cellWidth1)
        $fontSize2 = CalculateFontSize($承認者, $cellWidth2)
        $fontSize3 = CalculateFontSize($照査者, $cellWidth3)

        # 担当者のサイン欄
        $cell1.Range.Text = "$担当者`n$担当者日付"
        $cell1.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
        $cell1.Range.Font.Size = $fontSize1
        $cell1.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

        # 承認者のサイン欄
        $cell2.Range.Text = "$承認者`n$承認者日付"
        $cell2.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
        $cell2.Range.Font.Size = $fontSize2
        $cell2.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

        # 照査者のサイン欄
        $cell3.Range.Text = "$照査者`n$照査者日付"
        $cell3.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
        $cell3.Range.Font.Size = $fontSize3
        $cell3.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

        Write-Host "OUT: FillSignatures"
    }

    # ドキュメントを別名で保存するメソッド
    [void] SaveDocumentWithBackup() {
        Write-Host "IN: SaveDocumentWithBackup"
        $backupFilePath = Join-Path -Path $this.DocFilePath -ChildPath "backup_$($this.DocFileName)"
        $this.SaveAs($backupFilePath)
        Write-Host "OUT: SaveDocumentWithBackup"
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
$wordDoc.SetCustomProperty("CustomProperty1", "Value1")
$wordDoc.Check_Custom_Property()
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty21", "Value21", "C:\Users\y0927\Documents\GitHub\PS_Script\sample_temp.docx")

# 新しいインスタンスを作成してから操作を続行
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)


# サイン欄に名前と日付を配置
# $wordDoc.FillSignatures()

# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property("CustomProperty1")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property("CustomProperty2")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property("CustomProperty21")
Write-Host "Read Property Value: $propValue"

# カスタムプロパティを更新する
# $wordDoc.Update_Property("CustomProperty2", "UpdatedValue")

# カスタムプロパティを削除する
$wordDoc.Delete_Property("CustomProperty21")
$wordDoc.Check_Custom_Property()
# $wordDoc.Delete_Property("CustomProperty1")
# $wordDoc.Check_Custom_Property()

# ドキュメントを閉じる
$wordDoc.Close()