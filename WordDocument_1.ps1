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
        Write-Host "Entering Check_PC_Env"
        $envInfo = @{
            "PCName" = $env:COMPUTERNAME
            "PowerShellHome" = $env:PSHOME
            "IPAddress" = (Get-NetIPAddress -AddressFamily IPv4).IPAddress
            "DocFilePath" = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            "ScriptLibraryPath" = $this.ScriptRoot
        }
        $envInfo
        Write-Host "Exiting Check_PC_Env"
    }

    [void] Check_Word_Library() {
        Write-Host "Entering Check_Word_Library"
        $libraryPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
        if (Test-Path $libraryPath) {
            Write-Host "Word library found at $($libraryPath)"
        } else {
            Write-Host "Word library not found at $($libraryPath). Searching the entire system..."
            $libraryPath = Get-ChildItem -Path "C:\" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
            if ($libraryPath) {
                Write-Host "Word library found at $($libraryPath)"
            } else {
                Write-Host -ForegroundColor Red "Word library not found on this system."
                throw "Word library not found. Please install the required library."
            }
        }
        Write-Host "Exiting Check_Word_Library"
    }

    [void] Check_Custom_Property() {
        Write-Host "Entering Check_Custom_Property"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties
        $customPropsList = @()

        foreach ($prop in $customProps) {
            $customPropsList += $prop.Name
        }

        $customPropsList | Out-File -FilePath (Join-Path -Path $this.ScriptRoot -ChildPath "custom_properties.txt")
        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Check_Custom_Property"
    }

    [void] Create_Property([string]$propName, [string]$propValue) {
        Write-Host "Entering Create_Property with propName=$propName, propValue=$propValue"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        if ($null -eq $customProps) {
            Write-Host "customProps is null"
        } else {
            $customProps.Add($propName, $false, 4, $propValue)
            $doc.Save()
        }

        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Create_Property"
    }

    [string] Read_Property([string]$propName) {
        Write-Host "Entering Read_Property with propName=$propName"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties
        $propValue = $customProps.Item($propName).Value
        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Read_Property with propValue=$propValue"
        return $propValue
    }

    [void] Update_Property([string]$propName, [string]$propValue) {
        Write-Host "Entering Update_Property with propName=$propName, propValue=$propValue"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties
        $customProps.Item($propName).Value = $propValue
        $doc.Save()
        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Update_Property"
    }

    [void] Delete_Property([string]$propName) {
        Write-Host "Entering Delete_Property with propName=$propName"
        $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($docPath)
        $customProps = $doc.CustomDocumentProperties

        try {
            $prop = [System.__ComObject].InvokeMember("Item", "GetProperty", $null, $customProps, $propName)
            [System.__ComObject].InvokeMember("Delete", "Method", $null, $prop, $null)
        } catch {
            Write-Host -ForegroundColor Red "Property $($propName) not found."
        }

        $doc.Save()
        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Delete_Property"
    }

    [hashtable] Get_Properties([string]$PropertyType) {
        Write-Host "Entering Get_Properties with PropertyType=$PropertyType"
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
                    $pn = [System.__ComObject].InvokeMember("Item", "GetProperty", $null, $Properties, $p)
                    $value = [System.__ComObject].InvokeMember("Value", "GetProperty", $null, $pn, $null)
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch [System.Exception] {
                    Write-Host -ForegroundColor Blue "Value not found for $($p)"
                }
            }
        }

        if ($PropertyType -eq "Custom" -or $PropertyType -eq "Both") {
            $Properties = $doc.CustomDocumentProperties
            foreach ($p in $CustomPropertiesGroup) {
                try {
                    $pn = [System.__ComObject].InvokeMember("Item", "GetProperty", $null, $Properties, $p)
                    $value = [System.__ComObject].InvokeMember("Value", "GetProperty", $null, $pn, $null)
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch [System.Exception] {
                    Write-Host -ForegroundColor Blue "Value not found for $($p)"
                }
            }
        }

        $doc.Close()
        $word.Quit()
        Write-Host "Exiting Get_Properties"
        return $objHash
    }
}

# デバッグ用設定
$DocFileName = "技100-999.docx"
$DocFilePath1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$DocFilePath2 = "D:\Github\PS_Script"
$ScriptRoot = "C:\Users\y0927\Documents\GitHub\PS_Script"

# デバッグ環境に応じてパスを切り替える
$usePath1 = $true  # $true なら $DocFilePath1 を使用、$false なら $DocFilePath2 を使用

if ($usePath1) {
    $DocFilePath = $DocFilePath1
} else {
    $DocFilePath = $DocFilePath2
}

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)

# メソッドの呼び出し例
$wordDoc.Check_PC_Env()
$wordDoc.Check_Word_Library()
$wordDoc.Check_Custom_Property()
$wordDoc.Create_Property("NewProp", "NewValue")
$propValue = $wordDoc.Read_Property("NewProp")
$wordDoc.Update_Property("NewProp", "UpdatedValue")
$wordDoc.Delete_Property("NewProp")
$properties = $wordDoc.Get_Properties("Both")
