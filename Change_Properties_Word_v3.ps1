function Set-CustomProperty {
    param (
        [object]$doc,
        [string]$propName,
        [object]$propValue
    )

    try {
        $properties = $doc.CustomDocumentProperties
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
}

function Get-CustomProperty {
    param (
        [object]$doc,
        [string]$propName
    )

    try {
        $properties = $doc.CustomDocumentProperties
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

function Set-DocumentProperties {
    param (
        [object]$doc,
        [hashtable]$properties
    )

    foreach ($key in $properties.Keys) {
        try {
            $doc.PSObject.Properties[$key].Value = $properties[$key]
            Write-Host "Property '$key' set to '$($properties[$key])'"
        } catch {
            Write-Error "Failed to set property '$key': $_"
        }
    }
}

function ProcessDocument {
    param (
        [string]$filePath,
        [string]$approver,
        [bool]$approvalFlag,
        [string]$imagePath,
        [string]$iniFilePath
    )

    Import-InteropAssembly -iniFilePath $iniFilePath

    # スクリプト実行前に存在していたWordプロセスを取得
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

    # Wordアプリケーションを起動
    Write-Host "Wordアプリケーションを起動中..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    try {
        # ドキュメントを開く
        Write-Host "ドキュメントを開いています: $filePath"
        $doc = $word.Documents.Open($filePath)

        # 文書プロパティを表示
        Write-Host "現在の文書プロパティ:"
        Get-DocumentProperties -doc $doc

        # 承認者プロパティを設定
        Write-Host "承認者プロパティを設定中..."
        Set-CustomProperty -doc $doc -propName "Approver" -propValue $approver

        # 承認フラグプロパティを設定
        Write-Host "承認フラグプロパティを設定中..."
        Set-CustomProperty -doc $doc -propName "ApprovalFlag" -propValue ($approvalFlag ? "承認" : "未承認")

        # テーブル処理
        ProcessTable -doc $doc

        # 画像処理
        ProcessImage -doc $doc -imagePath $imagePath

        # ドキュメントを保存して閉じる
        Write-Host "ドキュメントを保存して閉じています..."
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

function Import-InteropAssembly {
    param (
        [string]$iniFilePath
    )

    $assemblyName = "Microsoft.Office.Interop.Word"
    Write-Host "Import-InteropAssemblyメソッドを実行中..."

    # INIファイルからアセンブリパスを読み込む
    if (Test-Path $iniFilePath) {
        $iniContent = Get-IniContent -iniFilePath $iniFilePath
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
                Set-Content -Path $iniFilePath -Value "$assemblyName=$assemblyPath"
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
                Set-Content -Path $iniFilePath -Value "$assemblyName=$assemblyPath"
                Add-Type -Path $assemblyPath
            }
        }
    }

    Write-Host "Import-InteropAssemblyメソッドが完了しました。"
}

function ProcessTable {
    param (
        [object]$doc
    )

    Write-Host "1つ目のテーブルを取得中..."
    try {
        $table = $doc.Tables.Item(1)
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

function ProcessImage {
    param (
        [object]$doc,
        [string]$imagePath
    )

    # 1つ目のセルを取得
    Write-Host "1つ目のセルを取得中..."
    $cell = $doc.Tables.Item(1).Cell(2, 6)
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
    foreach ($shape in $doc.Shapes) {
        if ($shape.Type -eq 3) { # 3 corresponds to wdInlineShapePicture
            $shape.Delete()
        }
    }
    Write-Host "Existing images deleted."

    # 新しい画像を挿入
    Write-Host "新しい画像を挿入中..."
    $shape = $doc.Shapes.AddPicture($imagePath, $false, $true, $imageLeft, $imageTop, $imageWidth, $imageHeight)
    Write-Host "New image inserted."

    # 画像のプロパティを変更
    Write-Host "画像のプロパティを変更中..."
    $shape.LockAspectRatio = $false
    $shape.Width = 50
    $shape.Height = 50
    Write-Host "Image properties modified."
}

# メインスクリプト
param (
    [string]$filePath,
    [string]$approver,
    [bool]$approvalFlag,
    [string]$imagePath,
    [string]$iniFilePath
)

ProcessDocument -filePath $filePath -approver $approver -approvalFlag $approvalFlag -imagePath $imagePath -iniFilePath $iniFilePath