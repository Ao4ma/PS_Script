# モジュールのインポート
using module ./PC_Class.psm1
using module ./Ini_Class.psm1

class Word {
    [object]$Application
    [object]$Document
    [hashtable]$DocumentProperties
    [MyPC]$PC
    [string]$IniFilePath

    Word([string]$filePath, [MyPC]$pc, [string]$iniFilePath) {
        $this.PC = $pc
        $this.IniFilePath = $iniFilePath
        if (-not $pc.IsLibraryConfigured) {
            Write-Error "Microsoft.Office.Interop.Word ライブラリが設定されていません。"
            return
        }

        $this.Application = New-Object -ComObject Word.Application
        $this.Application.Visible = $true
        $this.Document = $this.Application.Documents.Open($filePath)
        $this.DocumentProperties = $this.GetAllDocumentProperties()

        # プロパティをINIファイルに出力
        $iniFile = [IniFile]::new($this.IniFilePath)
        $iniFile.SetContent($this.DocumentProperties)

        # 5分後に自動解放
        $timer = New-Object Timers.Timer
        $timer.Interval = 300000  # 5分
        $timer.AutoReset = $false
        $timer.add_Elapsed({ $this.Close() })
        $timer.Start()
    }

    [bool]CheckLibraryConfigured() {
        return $this.PC.IsLibraryConfigured
    }

    [void]Close() {
        $this.Document.Close()
        $this.Application.Quit()
        Write-Host "Wordインスタンスが解放されました。"
        $global:WordInstances.Remove($this)
    }

    [void]SetDocumentProperty([string]$propertyName, [string]$newValue, [switch]$isNewCustomProperty) {
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

                # 新規カスタムプロパティの作成
                if ($isNewCustomProperty) {
                    if (-not $propertySet) {
                        try {
                            $Properties.Add($propertyName, $false, 4, $newValue)  # 4 corresponds to msoPropertyTypeString
                            Write-Host -ForegroundColor Green "Custom property '$propertyName' created and set to '$newValue'."
                            $propertySet = $true
                        } catch [System.Exception] {
                            Write-Host -ForegroundColor Red "Failed to create and set custom property '$propertyName'."
                        }
                    }
                } else {
                    if (-not $propertySet) {
                        Write-Host -ForegroundColor Red "Property '$propertyName' not found. It might be a typo."
                        return
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

    [string]GetDocumentProperty([string]$propertyName) {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $value = $null

        # ビルトインプロパティを取得
        $Properties = $this.Document.BuiltInDocumentProperties
        $value = Get-Property -Properties $Properties -propertyName $propertyName -binding $binding

        # カスタムプロパティを取得
        if ($null -eq $value) {
            $Properties = $this.Document.CustomDocumentProperties
            $value = Get-Property -Properties $Properties -propertyName $propertyName -binding $binding
        }

        return $value
    }

    [hashtable]GetAllDocumentProperties() {
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
    
            if ($null -eq $builtinProperties) {
                Write-Host "BuiltInDocumentProperties is null."
            } else {
                Write-Host "Type of BuiltInDocumentProperties: $($builtinProperties.GetType().FullName)"
                
                # ビルトインプロパティを取得
                foreach ($propertyName in $BuiltinPropertiesGroup) {
                    try {
                        $propertyValue = $builtinProperties.Item($propertyName).Value
                        $properties[$propertyName] = $propertyValue
                        Write-Host "$($propertyName): $propertyValue"
                    } catch {
                        Write-Host "Failed to get property '$propertyName': $_" -ForegroundColor Red
                    }
                }
            }
            Write-Host "END Standard Properties (ビルドインプロパティ) :"
    
            Write-Host "`nStart Custom Properties (カスタムプロパティ):"
            $customProperties = $this.Document.CustomDocumentProperties
            foreach ($propertyName in $CustomPropertiesGroup) {
                try {
                    $propertyValue = $customProperties.Item($propertyName).Value
                    $properties[$propertyName] = $propertyValue
                    Write-Host "$($propertyName): $propertyValue"
                } catch {
                    Write-Host "Failed to get property '$propertyName': $_" -ForegroundColor Red
                }
            }
            Write-Host "End Custom Properties (カスタムプロパティ):"
        } catch {
            Write-Error "プロパティの取得に失敗しました: $_"
        }
        return $properties
    }

    [void]AddCustomProperty([string]$propertyName, [string]$newValue) {
        try {
            $Properties = $this.Document.CustomDocumentProperties
            $Properties.Add($propertyName, $false, 4, $newValue)  # 4 corresponds to msoPropertyTypeString
            Write-Host -ForegroundColor Green "Custom property '$propertyName' created and set to '$newValue'."
            $this.Document.Save()
        } catch [System.Exception] {
            Write-Host -ForegroundColor Red "Failed to create and set custom property '$propertyName'."
        }
    }

    [void]RemoveCustomProperty([string]$propertyName) {
        try {
            $Properties = $this.Document.CustomDocumentProperties
            $existsBefore = $false
            $existsAfter = $false

            # 削除前に存在確認
            foreach ($prop in $Properties) {
                if ($prop.Name -eq $propertyName) {
                    $existsBefore = $true
                    break
                }
            }

            if ($existsBefore) {
                $Properties.Item($propertyName).Delete()
                Write-Host -ForegroundColor Green "Custom property '$propertyName' removed."
                $this.Document.Save()

                # 削除後に存在確認
                foreach ($prop in $Properties) {
                    if ($prop.Name -eq $propertyName) {
                        $existsAfter = $true
                        break
                    }
                }

                if (-not $existsAfter) {
                    Write-Host -ForegroundColor Green "Confirmed: Custom property '$propertyName' has been removed."
                } else {
                    Write-Host -ForegroundColor Red "Failed to remove custom property '$propertyName'."
                }
            } else {
                Write-Host -ForegroundColor Red "Custom property '$propertyName' does not exist."
            }
        } catch [System.Exception] {
            Write-Host -ForegroundColor Red "Failed to remove custom property '$propertyName'."
        }
    }

    [void]RecordTableCellInfo() {
        try {
            $tables = $this.Document.Tables
            foreach ($table in $tables) {
                foreach ($row in $table.Rows) {
                    foreach ($cell in $row.Cells) {
                        $cellText = $cell.Range.Text.Trim()
                        $left = $cell.Left
                        $top = $cell.Top
                        $right = $cell.Left + $cell.Width
                        $bottom = $cell.Top + $cell.Height
                        $centerX = ($left + $right) / 2
                        $centerY = ($top + $bottom) / 2

                        $propertyName = "CellInfo_$($cell.RowIndex)_$($cell.ColumnIndex)"
                        $propertyValue = "Text: $cellText, Left: $left, Top: $top, Right: $right, Bottom: $bottom, CenterX: $centerX, CenterY: $centerY"
                        $this.AddCustomProperty($propertyName, $propertyValue)
                    }
                }
            }
        } catch {
            Write-Host "An error occurred while recording table cell info: $_" -ForegroundColor Red
        }
    }

    [Boolean]SetProperty([object]$Properties, [string]$propertyName, [string]$newValue, [string]$binding) {
        try {
            $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $propertyName)
            [System.__ComObject].InvokeMember("Value", $binding::SetProperty, $null, $pn, $newValue)
            return $true
        } catch [System.Exception] {
            Write-Host -ForegroundColor Blue "Property '$propertyName' not found or cannot be set."
            return $false
        }
    }   

    [string]GetProperty([object]$Properties, [string]$propertyName, [string]$binding) {
        try {
            $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $propertyName)
            $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
            return $value
        } catch [System.Exception] {
            Write-Host -ForegroundColor Blue "Property '$propertyName' not found."
            return $null
        }
    }

    [void]GetProperties([object]$Properties, [array]$PropertyNames, [hashtable]$objHash, [string]$binding) {
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

    [String] GetIniContent ([string]$Path) {
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
}
