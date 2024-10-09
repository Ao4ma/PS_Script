# INIファイルのパス
$scriptPath = $MyInvocation.MyCommand.Path
$iniFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($scriptPath), "config_Change_Properties_Word.ini")

# INIファイルから設定を読み込む関数
function Get-IniContent {
    param (
        [string]$iniFilePath
    )
    $iniContent = @{}
    if (Test-Path $iniFilePath) {
        $lines = Get-Content $iniFilePath
        foreach ($line in $lines) {
            if ($line -match "^\s*([^=]+?)\s*=\s*(.*?)\s*$") {
                $iniContent[$matches[1].Trim()] = $matches[2].Trim()
            }
        }
    }
    return $iniContent
}

# INIファイルの内容を取得
$iniContent = Get-IniContent -iniFilePath $iniFilePath

# クラス定義
class WordDocumentProcessor {
    [string]$FilePath
    [string]$Approver
    [bool]$ApprovalFlag
    [string]$ImagePath

    WordDocumentProcessor([string]$filePath, [string]$approver, [bool]$approvalFlag, [string]$imagePath) {
        $this.FilePath = $filePath
        $this.Approver = $approver
        $this.ApprovalFlag = $approvalFlag
        $this.ImagePath = $imagePath
    }

    [void] ImportInteropAssembly() {
        $assemblyName = "Microsoft.Office.Interop.Word"
        $scriptPath = $MyInvocation.MyCommand.Path
        $iniFilePath = [System.IO.Path]::ChangeExtension($scriptPath, ".ini")

        # INIファイルからアセンブリパスを読み込む
        if (Test-Path $iniFilePath) {
            $assemblyPath = Get-Content $iniFilePath
        } else {
            $assemblyPath = $null
        }

        switch ($true) {
            ($assemblyPath -and (Test-Path $assemblyPath)) {
                Write-Output "$assemblyName is found in INI file. Using the existing assembly."
                Add-Type -Path $assemblyPath
            }
            ($null -eq $assemblyPath -or -not (Test-Path $assemblyPath)) {
                Write-Output "$assemblyName is not found in INI file or path does not exist. Searching in Windows directory..."

                # Windowsディレクトリ下をサーチ
                $assemblyPath = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" | Select-Object -First 1 -ExpandProperty FullName

                if ($assemblyPath) {
                    Write-Output "$assemblyName is found in Windows directory. Using the existing assembly."
                    Set-Content -Path $iniFilePath -Value $assemblyPath
                    Add-Type -Path $assemblyPath
                } else {
                    Write-Output "$assemblyName is not found in Windows directory. Installing from NuGet..."

                    # NuGetプロバイダーのインストール
                    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
                        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
                    }

                    # Microsoft.Office.Interop.Wordのインストール
                    Install-Package -Name "Microsoft.Office.Interop.Word" -Source "PSGallery" -Scope CurrentUser -Force
                    $assemblyPath = (Get-Package -Name "Microsoft.Office.Interop.Word" -Source "PSGallery").Source
                    Set-Content -Path $iniFilePath -Value $assemblyPath
                    Add-Type -Path $assemblyPath
                }
            }
        }
    }

    [void] ProcessDocument() {
        $this.ImportInteropAssembly()

        # スクリプト実行前に存在していたWordプロセスを取得
        $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

        # Wordアプリケーションを起動
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true

        try {
            # ドキュメントを開く
            $doc = $word.Documents.Open($this.FilePath)

            # 文書プロパティを読み取って表示する関数
            function Get-DocumentProperties {
                param (
                    [object]$doc
                )

                $properties = @{
                    DocumentTheme        = $doc.DocumentTheme
                    HasVBProject         = $doc.HasVBProject
                    OMathFontName        = $doc.OMathFontName
                    EncryptionProvider   = $doc.EncryptionProvider
                    UseMathDefaults      = $doc.UseMathDefaults
                    CurrentRsid          = $doc.CurrentRsid
                    DocID                = $doc.DocID
                    CompatibilityMode    = $doc.CompatibilityMode
                    CoAuthoring          = $doc.CoAuthoring
                    Broadcast            = $doc.Broadcast
                    ChartDataPointTrack  = $doc.ChartDataPointTrack
                    IsInAutosave         = $doc.IsInAutosave
                    WorkIdentity         = $doc.WorkIdentity
                    AutoSaveOn           = $doc.AutoSaveOn
                }

                foreach ($key in $properties.Keys) {
                    Write-Host "$($key): $($properties[$key])"
                }
            }

            # カスタムプロパティを設定する関数
            function Set-CustomProperty {
                param (
                    [object]$doc,
                    [string]$propName,
                    [object]$propValue
                )

                try {
                    $properties = $doc.CustomDocumentProperties
                } catch {
                    Write-Error "カスタムプロパティの取得に失敗しました: $_"
                    return
                }

                $property = $null

                # 既存のカスタムプロパティをチェック
                try {
                    $property = $properties.Item($propName)
                } catch {
                    # プロパティが存在しない場合は例外が発生するので無視
                }

                if ($property -ne $null) {
                    # 既存のプロパティを更新
                    $property.Value = $propValue
                } else {
                    # 新しいプロパティを追加
                    try {
                        $properties.Add($propName, $false, 4, $propValue) # 4はmsoPropertyTypeString
                    } catch {
                        Write-Error "カスタムプロパティの追加に失敗しました: $_"
                    }
                }
            }

            # 文書プロパティを表示
            Write-Host "現在の文書プロパティ:"
            Get-DocumentProperties -doc $doc

            # 承認者プロパティを設定
            Set-CustomProperty -doc $doc -propName "承認者" -propValue $this.Approver

            # 承認フラグプロパティを設定
            Set-CustomProperty -doc $doc -propName "承認フラグ" -propValue ([string]$this.ApprovalFlag)

            # 1つ目のテーブルを取得
            $table = $doc.Tables.Item(1)
            Write-Host "First table retrieved."

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

            # 1つ目のセルを取得
            $cell = $table.Cell(2, 6)
            Write-Host "Cell (2, 6) retrieved."

            # セルの座標とサイズを取得
            $left = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdHorizontalPositionRelativeToPage)
            $top = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
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
            foreach ($shape in $doc.Shapes) {
                if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdInlineShapeType]::wdInlineShapePicture) {
                    $shape.Delete()
                }
            }
            Write-Host "Existing images deleted."

            # 新しい画像を挿入
            $shape = $doc.Shapes.AddPicture($this.ImagePath, $false, $true, $imageLeft, $imageTop, $imageWidth, $imageHeight)
            Write-Host "New image inserted."

            # 画像のプロパティを変更
            $shape.LockAspectRatio = $false
            $shape.Width = 50
            $shape.Height = 50
            Write-Host "Image properties modified."

            # ドキュメントを保存して閉じる
            $doc.Save()
            $doc.Close()
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

            # Wordアプリケーションを終了
            try {
                $word.Quit()
            } catch {
                Write-Warning "Wordアプリケーションの終了に失敗しました: $($_.Exception.Message)"
            }
        }

        Write-Host "カスタムプロパティが設定されました。"
    }
}

# INIファイルから設定を読み込む
$filePath = $iniContent["FilePath"]
$approver = $iniContent["Approver"]
$approvalFlag = $false
if ($iniContent.ContainsKey("ApprovalFlag")) {
    $approvalFlag = [bool]::Parse($iniContent["ApprovalFlag"])
}
$imagePath = $iniContent["ImagePath"]

# クラスのインスタンスを作成して処理を実行
$processor = [WordDocumentProcessor]::new($filePath, $approver, $approvalFlag, $imagePath)
$processor.ProcessDocument()