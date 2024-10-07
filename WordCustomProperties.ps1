# PowerShellスクリプト: WordCustomProperties.ps1

param (
    [string]$filePath = "C:\Users\y0927\Documents\GitHub\PS_Script\work\技100-999.docx",
    [string]$承認者 = "青島",
    [bool]$承認フラグ = $true
)

# スクリプト実行前に存在していたWordプロセスを取得
$existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

# Wordアプリケーションを起動
$word = New-Object -ComObject Word.Application
$word.Visible = $true

try {
    # ドキュメントを開く
    $doc = $word.Documents.Open($filePath)

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

        $properties = $doc.CustomDocumentProperties
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
            $properties.Add($propName, $false, 4, $propValue) # 4はmsoPropertyTypeString
        }
    }

    # 文書プロパティを表示
    Write-Host "現在の文書プロパティ:"
    Get-DocumentProperties -doc $doc


    # 承認者プロパティを設定
    Set-CustomProperty -doc $doc -propName "承認者" -propValue $承認者

    # 承認フラグプロパティを設定
    Set-CustomProperty -doc $doc -propName "承認フラグ" -propValue ([string]$承認フラグ)

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