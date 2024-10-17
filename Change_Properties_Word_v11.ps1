# スクリプトのフォルダパスを取得
$scriptFolderPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$iniFilePath = Join-Path -Path $scriptFolderPath -ChildPath "config_Change_Properties_Word.ini"

# IniFileクラスを読み込む
. "C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\Ini_Class.ps1"

# PCクラスを読み込む
. "C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\PC_Class.ps1"

# Wordクラスを読み込む
. "C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\Word_Class.ps1"

function ProcessDocument {
    param (
        [PC]$pc,
        [string]$filePath
    )

    if (-not (Test-Path $filePath)) {
        Write-Error "ファイルパスが無効です: $filePath"
        return
    }

    # スクリプト実行前に存在していたWordプロセスを取得
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

    # Wordアプリケーションを起動
    Write-Host "Wordアプリケーションを起動中..."
    $word = [Word]::new($filePath, $pc)

    try {
        # 文書プロパティを表示
        Write-Host "現在の文書プロパティ:"
        foreach ($property in $word.DocumentProperties) {
            Write-Host "$($property.Key): $($property.Value)"
        }

        # 承認者プロパティを設定
        # Write-Host "承認者プロパティを設定中..."
        # $wordInstanceData.SetCustomProperty("Approver", $approver)

        # 承認フラグプロパティを設定
        # Write-Host "承認フラグプロパティを設定中..."
        # $wordInstanceData.SetCustomProperty("ApprovalFlag", ($approvalFlag ? "承認" : "未承認"))

        # カスタムプロパティの追加例
        $word.AddCustomProperty("NewCustomProperty", "NewValue")

        # カスタムプロパティの削除例
        $word.RemoveCustomProperty("OldCustomProperty")

        # テーブルセル情報の記録
        $word.RecordTableCellInfo()

        # 変更後の文書プロパティを表示
        Write-Host "変更後の文書プロパティ:"
        foreach ($property in $word.DocumentProperties) {
            Write-Host "$($property.Key): $($property.Value)"
        }
    } finally {
        # Wordアプリケーションを閉じる
        $word.Close()

        # スクリプト実行後に新たに起動されたWordプロセスを終了
        $newWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
        foreach ($process in $newWordProcesses) {
            if ($existingWordProcesses -notcontains $process) {
                Stop-Process -Id $process.Id -Force
            }
        }
    }
}

# IniFileクラスのインスタンスを作成
$ini = [IniFile]::new($iniFilePath)

# PCクラスのインスタンスを作成
$pc = New-Object -TypeName PSObject -Property @{
    IsLibraryConfigured = $true
}

# ドキュメントのパスを取得
$filePath = $ini.GetValue("Settings", "FilePath")

# ドキュメントを処理
ProcessDocument -pc $pc -filePath $filePath