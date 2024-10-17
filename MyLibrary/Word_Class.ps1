class Word {
    [object]$Application
    [object]$Document
    [hashtable]$DocumentProperties
    [object]$PC

    Word([string]$filePath, [object]$pc) {
        $this.PC = $pc
        if (-not $this.PC.IsLibraryConfigured) {
            Write-Error "Microsoft.Office.Interop.Word ライブラリが設定されていません。"
            return
        }

        $this.Application = New-Object -ComObject Word.Application
        $this.Application.Visible = $true
        $this.Document = $this.Application.Documents.Open($filePath)
        $this.DocumentProperties = $this.GetDocumentProperties()
    }

    [void]Close() {
        $this.Document.Close()
        $this.Application.Quit()
    }

    [void]SetDocumentProperty([string]$propertyName, [string]$newValue) {
        try {
            $binding = "System.Reflection.BindingFlags" -as [type]
            $propertySet = $false

            # ビルトインプロパティを設定
            $Properties = $this.Document.BuiltInDocumentProperties
            $propertySet = Set-Property -Properties $Properties -propertyName $propertyName -newValue $newValue -binding $binding

            # カスタムプロパティを設定
            if (-not $propertySet) {
                $Properties = $this.Document.CustomDocumentProperties
                $propertySet = Set-Property -Properties $Properties -propertyName $propertyName -newValue $newValue -binding $binding

                # カスタムプロパティが存在しない場合、新規作成
                if (-not $propertySet) {
                    try {
                        $Properties.Add($propertyName, $false, 4, $newValue)  # 4 corresponds to msoPropertyTypeString
                        Write-Host -ForegroundColor Green "Custom property '$propertyName' created and set to '$newValue'."
                        $propertySet = $true
                    } catch [System.Exception] {
                        Write-Host -ForegroundColor Red "Failed to create and set custom property '$propertyName'."
                    }
                }
            }

            # ドキュメントを保存して閉じる
            $this.Document.Save()
            if ($propertySet) {
                Write-Host "Property '$propertyName' set to '$newValue'." -ForegroundColor Green
            } else {
                Write-Host "Failed to set property '$propertyName'." -ForegroundColor Red
            }
        } catch {
            Write-Host "An error occurred: $_" -ForegroundColor Red
        }
    }

    [hashtable]GetDocumentProperties() {
        $properties = @{}
        $binding = "System.Reflection.BindingFlags" -as [type]

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

        $CustomPropertiesGroup = @("batter", "yamada", "Path")

        try {
            Write-Host "Start Standard Properties (ビルドインプロパティ):"
            $builtinProperties = $this.Document.BuiltInDocumentProperties
            Get-Properties -Properties $builtinProperties -PropertyNames $BuiltinPropertiesGroup -objHash $properties -binding $binding
            Write-Host "END Standard Properties (ビルドインプロパティ) :"

            Write-Host "`nStart Custom Properties (カスタムプロパティ):"
            $customProperties = $this.Document.CustomDocumentProperties
            Get-Properties -Properties $customProperties -PropertyNames $CustomPropertiesGroup -objHash $properties -binding $binding
            Write-Host "End Custom Properties (カスタムプロパティ):"
        } catch {
            Write-Error "プロパティの取得に失敗しました: $_"
        }
        return $properties
    }
}

function Set-Property {
    param (
        [object]$Properties,
        [string]$propertyName,
        [string]$newValue,
        [string]$binding
    )

    try {
        $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $propertyName)
        [System.__ComObject].InvokeMember("Value", $binding::SetProperty, $null, $pn, $newValue)
        return $true
    } catch [System.Exception] {
        Write-Host -ForegroundColor Blue "Property '$propertyName' not found or cannot be set."
        return $false
    }
}

function Get-Properties {
    param (
        [object]$Properties,
        [array]$PropertyNames,
        [hashtable]$objHash,
        [string]$binding
    )

    foreach ($p in $PropertyNames) {
        try {
            $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $p)
            $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
            $objHash[$p] = $value
        } catch [System.Exception] {
            Write-Host -ForegroundColor Blue "Value not found for $p"
        }
    }
}

function Get-IniContent {
    param (
        [string]$Path
    )

    $iniContent = @{}
    $currentSection = ""

    foreach ($line in Get-Content -Path $Path) {
        if ($line -match "^\[(.+)\]$") {
            $currentSection = $matches[1]
            $iniContent[$currentSection] = @{}
        } elseif ($line -match "^(.+?)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $iniContent[$currentSection][$key] = $value
        }
    }

    return $iniContent
}