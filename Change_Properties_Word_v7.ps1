param (
    [string]$filePath,
    [string]$approver,
    [bool]$approvalFlag,
    [string]$imagePath,
    [string]$iniFilePath
)

class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress

    PC([string]$name, [string]$iniFilePath) {
        $this.Name = $name
        $this.IPAddress = $this.GetIPAddress()
        $this.MACAddress = $this.GetMACAddress()
        $this.Import_InteropAssembly($iniFilePath)
    }

    [void]DisplayInfo() {
        Write-Host "PC Name: $($this.Name)"
        Write-Host "IP Address: $($this.IPAddress)"
        Write-Host "MAC Address: $($this.MACAddress)"
    }

    [string]GetIPAddress() {
        $ipConfig = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" }
        return $ipConfig.IPAddress
    }

    [string]GetMACAddress() {
        $macConfig = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        return $macConfig.MacAddress
    }

    [string]GetScriptDirectory() {
        return Split-Path -Parent -Path $PSCommandPath
    }

    [void]ChangeToScriptDirectory() {
        $scriptDir = $this.GetScriptDirectory()
        Set-Location -Path $scriptDir
        Write-Host "Changed directory to script location: $scriptDir"
    }

    [hashtable]GetIniContent([string]$iniFilePath) {
        $iniContent = @{}
        if (Test-Path $iniFilePath) {
            $lines = Get-Content -Path $iniFilePath
            foreach ($line in $lines) {
                if ($line -match "^\s*`"([^\`"]+?)`"\s*,\s*`"([^\`"]+?)`"\s*$") {
                    $key = $matches[1]
                    $value = $matches[2]
                    $iniContent[$key] = $value
                }
            }
        }
        return $iniContent
    }

    [void]SavePCInfoToIni([string]$iniFilePath) {
        $iniContent = $this.GetIniContent($iniFilePath)
        $iniContent[$this.Name] = "$($this.IPAddress),$($this.MACAddress),$($this.GetScriptDirectory())"
        $iniLines = @()
        foreach ($key in $iniContent.Keys) {
            $iniLines += "`"$key`",`"$($iniContent[$key])`""
        }
        Set-Content -Path $iniFilePath -Value $iniLines
    }

    [void]LoadPCInfoFromIni([string]$iniFilePath) {
        $iniContent = $this.GetIniContent($iniFilePath)
        if ($iniContent.ContainsKey($this.Name)) {
            $info = $iniContent[$this.Name] -split ","
            $this.IPAddress = $info[0]
            $this.MACAddress = $info[1]
            $scriptDir = $info[2]
            Set-Location -Path $scriptDir
            Write-Host "Loaded PC info from INI file: $($this.Name), $($this.IPAddress), $($this.MACAddress), $scriptDir"
        } else {
            Write-Host "PC info not found in INI file for: $($this.Name)"
        }
    }

    [void]Import_InteropAssembly([string]$iniFilePath) {
        $assemblyName = "Microsoft.Office.Interop.Word"
        Write-Host "Import_InteropAssemblyメソッドを実行中..."

        # INIファイルからアセンブリパスを読み込む
        if (Test-Path $iniFilePath) {
            $iniContent = $this.GetIniContent($iniFilePath)
            $assemblyPath = $iniContent[$assemblyName]
        } else {
            $assemblyPath = $null
        }

        switch ($true) {
            ($assemblyPath -and (Test-Path $assemblyPath)) {
                Write-Host "$assemblyName is found in INI file. Using the existing assembly."
                Add-Type -Path $assemblyPath
            }
            ($null -eq $assemblyPath -or -not (Test-Path $assemblyPath)) {
                Write-Host "$assemblyName is not found in INI file or path does not exist. Searching in Windows directory..."

                # Windowsディレクトリ下をサーチ
                $assemblyPath = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" | Select-Object -First 1 -ExpandProperty FullName

                if ($assemblyPath) {
                    Write-Host "$assemblyName is found in Windows directory. Using the existing assembly."
                    if ($iniFilePath) {
                        Set-Content -Path $iniFilePath -Value "$assemblyName=$assemblyPath"
                    }
                    Add-Type -Path $assemblyPath
                } else {
                    Write-Host "$assemblyName is not found in Windows directory. Installing from NuGet..."

                    # NuGetプロバイダーのインストール
                    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
                        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
                    }

                    # Microsoft.Office.Interop.Wordのインストール
                    Install-Package -Name "Microsoft.Office.Interop.Word" -Source "PSGallery" -Scope CurrentUser -Force
                    $assemblyPath = (Get-Package -Name "Microsoft.Office.Interop.Word" -Source "PSGallery").Source
                    if ($iniFilePath) {
                        Set-Content -Path $iniFilePath -Value "$assemblyName=$assemblyPath"
                    }
                    Add-Type -Path $assemblyPath
                }
            }
        }

        Write-Host "Import_InteropAssemblyメソッドが完了しました。"
    }
}
class Word {
    [object]$Application
    [object]$Document
    [hashtable]$DocumentProperties
    [hashtable]$CustomProperties

    Word([string]$filePath) {
        $this.Application = New-Object -ComObject Word.Application
        $this.Application.Visible = $true
        $this.Document = $this.Application.Documents.Open($filePath)
        $this.DocumentProperties = $this.GetDocumentProperties($filePath)
        $this.CustomProperties = $this.GetCustomProperties()
    }

    [void]Close() {
        $this.Document.Close()
        $this.Application.Quit()
    }

    [void]SetCustomProperty([string]$propName, [object]$propValue) {
        try {
            $properties = $this.Document.CustomDocumentProperties
            if ($null -eq $properties) {
                Write-Error "カスタムプロパティが見つかりませんでした。"
                return
            }
        } catch {
            Write-Error "カスタムプロパティの取得に失敗しました: $_"
            return
        }

        # カスタムプロパティの一覧を表示（デバッグ用）
        foreach ($property in $properties) {
            Write-Host "Property Name: $($property.Name), Property Value: $($property.Value)"
        }

        # 既存のカスタムプロパティをチェック
        try {
            $property = $properties.Item($propName)
            Write-Host "Property found: $($property.Name), Value: $($property.Value)"
        } catch {
            # プロパティが存在しない場合は例外が発生するので無視
            Write-Host "Property '$propName' not found."
            $property = $null
        }

        if ($null -ne $property) {
            if ($null -eq $propValue) {
                # プロパティを削除
                try {
                    Write-Host "Deleting custom property: $propName"
                    $properties.Remove($propName)
                } catch {
                    Write-Error "カスタムプロパティの削除に失敗しました: $_"
                }
            } else {
                # 既存のプロパティを更新
                Write-Host "Updating custom property: $propName = $propValue"
                $property.Value = $propValue
            }
        } else {
            if ($null -ne $propValue) {
                # 新しいプロパティを追加
                try {
                    Write-Host "Adding new custom property: $propName = $propValue"
                    $properties.Add($propName, $false, 4, $propValue) # 4はmsoPropertyTypeString
                } catch {
                    Write-Error "カスタムプロパティの追加に失敗しました: $_"
                }
            }
        }

        # クラス変数を更新
        $this.CustomProperties = $this.GetCustomProperties()
    }

    [object]GetCustomProperty([string]$propName) {
        try {
            $properties = $this.Document.CustomDocumentProperties
            if ($null -eq $properties) {
                Write-Error "カスタムプロパティが見つかりませんでした。"
                return $null
            }
        } catch {
            Write-Error "カスタムプロパティの取得に失敗しました: $_"
            return $null
        }

        # 既存のカスタムプロパティをチェック
        try {
            $property = $properties.Item($propName)
            Write-Host "Property found: $($property.Name), Value: $($property.Value)"
            return $property.Value
        } catch {
            # プロパティが存在しない場合は例外が発生するので無視
            Write-Host "Property '$propName' not found."
            return $null
        }
    }

    [hashtable]GetCustomProperties() {
        $properties = @{}
        try {
            $this.CustomProperties = $this.Document.CustomDocumentProperties
            foreach ($property in $this.CustomProperties) {
                $properties[$property.Name] = $property.Value
            }
        } catch {
            Write-Error "カスタムプロパティの取得に失敗しました: $_"
        }
        return $properties
    }

    [hashtable]GetDocumentProperties([string]$filePath) {
        $properties = @{}
        try {
            $itemProperties = Get-ItemProperty -Path $filePath
            $properties["Author"] = $itemProperties.Author
            $properties["Title"] = $itemProperties.Title
            $properties["Subject"] = $itemProperties.Subject
            $properties["Keywords"] = $itemProperties.Keywords
            $properties["Category"] = $itemProperties.Category
            $properties["Comments"] = $itemProperties.Comments
            $properties["RevisionNumber"] = $itemProperties.RevisionNumber
            $properties["Version"] = $itemProperties.Version
        } catch {
            Write-Warning "Failed to get document properties: $_"
        }
        return $properties
    }

    [void]SetDocumentProperties([hashtable]$properties) {
        foreach ($key in $properties.Keys) {
            try {
                $this.Document.BuiltInDocumentProperties($key).Value = $properties[$key]
                Write-Host "Property '$key' set to '$($properties[$key])'"
            } catch {
                Write-Error "Failed to set property '$key': $_"
            }
        }

        # クラス変数を更新
        $this.DocumentProperties = $this.GetDocumentProperties($this.Document.FullName)
    }

    [void]ProcessTable() {
        Write-Host "1つ目のテーブルを取得中..."
        try {
            $table = $this.Document.Tables.Item(1)
            Write-Host "First table retrieved."
        } catch {
            Write-Error "テーブルの取得に失敗しました: $_"
            return
        }

        # テーブルのプロパティを取得
        $rows = $table.Rows.Count
        $columns = $table.Columns.Count
        Write-Host "Table properties retrieved: Rows=$rows, Columns=$columns"

        # 各セルの情報を取得
        foreach ($row in 1..$rows) {
            foreach ($col in 1..$columns) {
                try {
                    $cell = $table.Cell($row, $col)
                    $cellText = $cell.Range.Text
                    Write-Host "Row: $row, Column: $col, Text: $cellText"
                } catch {
                    Write-Host "Row: $row, Column: $col, Text: (cell not found)" -Foreground Red
                }
            }
        }
    }

    [void]ProcessImage([string]$imagePath) {
        # 1つ目のセルを取得
        Write-Host "1つ目のセルを取得中..."
        $cell = $this.Document.Tables.Item(1).Cell(2, 6)
        Write-Host "Cell (2, 6) retrieved."

        # セルの座標とサイズを取得
        $left = $cell.Range.Information(1) # 1 corresponds to wdHorizontalPositionRelativeToPage
        $top = $cell.Range.Information(2) # 2 corresponds to wdVerticalPositionRelativeToPage
        $width = $cell.Width
        $height = $cell.Height
        Write-Host "Cell coordinates and size retrieved: Left=$left, Top=$top, Width=$width, Height=$height"

        # 画像のサイズを設定
        $imageWidth = 50
        $imageHeight = 50

        # 画像の中央位置を計算
        $imageLeft = $left + ($width - $imageWidth) / 2
        $imageTop = $top + ($height - $imageHeight) / 2

        # 既存の画像を削除（もしあれば）
        Write-Host "既存の画像を削除中..."
        foreach ($shape in $this.Document.Shapes) {
            if ($shape.Type -eq 3) { # 3 corresponds to wdInlineShapePicture
                $shape.Delete()
            }
        }
        Write-Host "Existing images deleted."

        # 新しい画像を挿入
        Write-Host "新しい画像を挿入中..."
        $shape = $this.Document.Shapes.AddPicture($imagePath, $false, $true, $imageLeft, $imageTop, $imageWidth, $imageHeight)
        Write-Host "New image inserted."

        # 画像のプロパティを変更
        Write-Host "画像のプロパティを変更中..."
        $shape.LockAspectRatio = $false
        $shape.Width = 50
        $shape.Height = 50
        Write-Host "Image properties modified."
    }
}

function ProcessDocument {
    param (
        [PC]$pc,
        [string]$filePath,
        [string]$approver,
        [bool]$approvalFlag,
        [string]$imagePath,
        [string]$iniFilePath
    )

    if (-not (Test-Path $filePath)) {
        Write-Error "ファイルパスが無効です: $filePath"
        return
    }

    # スクリプト実行前に存在していたWordプロセスを取得
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

    # Wordアプリケーションを起動
    Write-Host "Wordアプリケーションを起動中..."
    $word = [Word]::new($filePath)

    try {
        # 文書プロパティを表示
        Write-Host "現在の文書プロパティ:"
        foreach ($key in $word.DocumentProperties.Keys) {
            Write-Host "$($key): $($word.DocumentProperties[$key])"
        }

        # 承認者プロパティを設定
        Write-Host "承認者プロパティを設定中..."
        $word.SetCustomProperty("Approver", $approver)

        # 承認フラグプロパティを設定
        Write-Host "承認フラグプロパティを設定中..."
        $word.SetCustomProperty("ApprovalFlag", ($approvalFlag ? "承認" : "未承認"))

        # テーブル処理
        $word.ProcessTable()

        # 画像処理
        $word.ProcessImage($imagePath)

        # ドキュメントを保存して閉じる
        Write-Host "ドキュメントを保存して閉じています..."
        $word.Document.Save()
        $word.Close()
    } catch {
        Write-Error "エラーが発生しました: $_"
    } finally {
        # スクリプト実行後に存在するWordプロセスを取得
        $allWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

        # スクリプト実行前に存在していたプロセスを除外して終了
        $newWordProcesses = $allWordProcesses | Where-Object { $_.Id -notin $existingWordProcesses.Id }
        foreach ($proc in $newWordProcesses) {
            try {
                $proc.Kill()
            } catch {
                Write-Warning "プロセスの終了に失敗しました: $($_.Exception.Message)"
            }
        }
    }

    Write-Host "カスタムプロパティが設定されました。"
}

# PCクラスのインスタンスを作成し、スクリプトのあるフォルダに移動
$PcName = "DELLD033"
$pc = [PC]::new($PcName, $iniFilePath)
$pc.ChangeToScriptDirectory()

# デバッグ用変数
$filePath = Join-Path -Path (Get-Location) -ChildPath "技100-999.docx"
$approver = "大谷"
$approvalFlag = $true
$imagePath = Join-Path -Path (Get-Location) -ChildPath "社長印.tif"
$iniFilePath = Join-Path -Path (Get-Location) -ChildPath "config_Change_Properties_Word.ini"

# メインスクリプト
ProcessDocument -pc $pc -filePath $filePath -approver $approver -approvalFlag $approvalFlag -imagePath $imagePath -iniFilePath $iniFilePath