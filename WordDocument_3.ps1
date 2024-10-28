class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
    }

    [void] Check_PC_Env() {
        Write-Host "IN: Check_PC_Env"
        $envInfo = @{
            "PCName" = $env:COMPUTERNAME
            "PowerShellHome" = $env:PSHOME
            "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
            "DocFilePath" = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            "ScriptLibraryPath" = $this.ScriptRoot
        }
        $envInfo
        Write-Host "OUT: Check_PC_Env"
    }

    [void] Check_Word_Library() {
        Write-Host "IN: Check_Word_Library"
        $libraryPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
        if (Test-Path $libraryPath) {
            Write-Host "Microsoft.Office.Interop.Word ライブラリが見つかりました。"
        } else {
            Write-Host -ForegroundColor Red "Microsoft.Office.Interop.Word ライブラリが見つかりません。"
        }
        Write-Host "OUT: Check_Word_Library"
    }

    [void] Create_Property([string]$propName, [string]$propValue) {
        Write-Host "IN: Create_Property"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        # 値を確認
        Write-Host "CustomDocumentProperties: $customProps"
        Write-Host "PropName: $propName"
        Write-Host "PropValue: $propValue"

        if ($null -eq $customProps) {
            Write-Host -ForegroundColor Red "CustomDocumentProperties is null. Cannot add property."
            $doc.Close()
            $word.Quit()
            Write-Host "OUT: Create_Property"
            return
        }

        try {
            $customProps.Add($propName, $false, 4, $propValue)
        } catch {
            Write-Host -ForegroundColor Red "Failed to add property '$propName': $_"
            try {
                $prop = $customProps.Item($propName)
                $prop.Delete()
                $customProps.Add($propName, $false, 4, $propValue)
            } catch {
                Write-Host -ForegroundColor Red "Failed to delete and re-add property '$propName': $_"
            }
        }

        $doc.Save()
        $doc.Close()
        $word.Quit()
        Write-Host "OUT: Create_Property"
    }

    [string] Read_Property([string]$propName) {
        Write-Host "IN: Read_Property"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        if ($null -eq $customProps) {
            Write-Host -ForegroundColor Red "CustomDocumentProperties is null. Cannot read property."
            $doc.Close()
            $word.Quit()
            Write-Host "OUT: Read_Property"
            return $null
        }

        try {
            $prop = $customProps.Item($propName)
            if ($null -eq $prop) {
                Write-Host -ForegroundColor Red "Property '$propName' not found."
                $doc.Close()
                $word.Quit()
                Write-Host "OUT: Read_Property"
                return $null
            }
            $propValue = $prop.Value
        } catch {
            Write-Host -ForegroundColor Red "Failed to read property '$propName': $_"
            $doc.Close()
            $word.Quit()
            Write-Host "OUT: Read_Property"
            return $null
        }

        $doc.Close()
        $word.Quit()
        Write-Host "OUT: Read_Property"
        return $propValue
    }

    [void] Update_Property([string]$propName, [string]$propValue) {
        Write-Host "IN: Update_Property"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        if ($null -eq $customProps) {
            Write-Host -ForegroundColor Red "CustomDocumentProperties is null. Cannot update property."
            $doc.Close()
            $word.Quit()
            Write-Host "OUT: Update_Property"
            return
        }

        try {
            $prop = $customProps.Item($propName)
            if ($null -eq $prop) {
                Write-Host -ForegroundColor Red "Property '$propName' not found."
                $doc.Close()
                $word.Quit()
                Write-Host "OUT: Update_Property"
                return
            }
            $prop.Value = $propValue
        } catch {
            Write-Host -ForegroundColor Red "Failed to update property '$propName': $_"
        }

        $doc.Save()
        $doc.Close()
        $word.Quit()
        Write-Host "OUT: Update_Property"
    }

    [void] Delete_Property([string]$propName) {
        Write-Host "IN: Delete_Property"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        if ($null -eq $customProps) {
            Write-Host -ForegroundColor Red "CustomDocumentProperties is null. Cannot delete property."
            $doc.Close()
            $word.Quit()
            Write-Host "OUT: Delete_Property"
            return
        }

        try {
            $prop = $customProps.Item($propName)
            $prop.Delete()
        } catch {
            Write-Host -ForegroundColor Red "Property $($propName) not found."
        }

        $doc.Save()
        $doc.Close()
        $word.Quit()
        Write-Host "OUT: Delete_Property"
    }

    [hashtable] Get_Properties([string]$PropertyType) {
        Write-Host "IN: Get_Properties"
        $BuiltinPropertiesGroup = @(
            "Title", "Subject", "Author", "Keywords", "Comments", "Template", "Last Author", 
            "Revision Number", "Application Name", "Last Print Date", "Creation Date", 
            "Last Save Time", "Total Editing Time", "Number of Pages", "Number of Words", 
            "Number of Characters", "Security", "Category", "Format", "Manager", "Company", 
            "Number of Bytes", "Number of Lines", "Number of Paragraphs", "Number of Slides", 
            "Number of Notes", "Number of Hidden Slides", "Number of Multimedia Clips", 
            "Hyperlink Base", "Number of Characters (with spaces)", "Content Type", 
            "Content Status", "Language", "Document Version"
        )

        $CustomPropertiesGroup = @("batter", "yamada", "Path") # これは例

        $objHash = @{}
        $foundProperties = @()

        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)

        if ($PropertyType -eq "Builtin" -or $PropertyType -eq "Both") {
            $Properties = $doc.BuiltInDocumentProperties
            foreach ($p in $BuiltinPropertiesGroup) {
                try {
                    $pn = $Properties.Item($p)
                    $value = $pn.Value
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch {
                    Write-Host -ForegroundColor Blue "Value not found for $($p)"
                }
            }
        }

        if ($PropertyType -eq "Custom" -or $PropertyType -eq "Both") {
            $Properties = $doc.CustomDocumentProperties
            foreach ($p in $CustomPropertiesGroup) {
                try {
                    $pn = $Properties.Item($p)
                    $value = $pn.Value
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch {
                    Write-Host -ForegroundColor Blue "Value not found for $($p)"
                }
            }
        }

        # ドキュメントを保存せずに閉じる
        [ref]$SaveOption = "Microsoft.Office.Interop.Word.WdSaveOptions" -as [type]
        $doc.Close($SaveOption)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Properties) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        Remove-Variable -Name doc, Properties

        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Remove-Variable -Name word
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()

        Write-Host "OUT: Get_Properties"
        return $objHash
    }

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
$wordDoc.Create_Property("NewProp", "NewValue")
$propValue = $wordDoc.Read_Property("NewProp")
$wordDoc.Update_Property("NewProp", "UpdatedValue")
$wordDoc.Delete_Property("NewProp")
$properties = $wordDoc.Get_Properties("Both")