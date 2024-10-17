# スクリプトのフォルダパスを取得
$scriptFolderPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$iniFilePath = Join-Path -Path $scriptFolderPath -ChildPath "config_Change_Properties_Word.ini"

class Main {
    [PC]$pc
    [IniFile]$ini
    [string]$filePath

    Main() {
        # PCクラスのインスタンスを作成
        $this.pc = [PC]::new()

        # スクリプト実行フォルダ、ログフォルダ、スクリプトフォルダを設定
        $this.pc.SetScriptFolder($scriptFolderPath)
        $this.pc.SetLogFolder("$scriptFolderPath\Logs")
        $this.pc.SetScriptFolder($scriptFolderPath)

        # IniFileクラスのインスタンスを作成
        $this.ini = [IniFile]::new($iniFilePath)

        # PC情報を調べて、iniファイルに記録
        $this.RecordPCInfo()

        # ドキュメントのパスを取得
        $this.filePath = $this.ini.GetValue("Settings", "FilePath")

        # ドキュメントを処理
        $this.ProcessDocument()
    }

    [void]RecordPCInfo() {
        $pcInfo = @{
            "OS" = (Get-WmiObject -Class Win32_OperatingSystem).Caption
            "UserName" = $env:USERNAME
            "MachineName" = $env:COMPUTERNAME
        }

        foreach ($key in $pcInfo.Keys) {
            $this.ini.SetValue("PCInfo", $key, $pcInfo[$key])
        }
    }

    [void]ProcessDocument() {
        if (-not (Test-Path $this.filePath)) {
            Write-Error "ファイルパスが無効です: $this.filePath"
            return
        }

        # スクリプト実行前に存在していたWordプロセスを取得
        $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

        # Wordアプリケーションを起動
        Write-Host "Wordアプリケーションを起動中..."
        $word = [Word]::new($this.filePath, $this.pc)

        try {
            # 文書プロパティを表示
            Write-Host "現在の文書プロパティ:"
            foreach ($property in $word.DocumentProperties) {
                Write-Host "$($property.Key): $($property.Value)"
            }

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

            # PCインスタンスへ通知してインスタンスの管理を終了
            $this.pc.NotifyInstanceClosed($word)
        }
    }
}

# PCクラスのインスタンスを作成
$pc = [PC]::new()

# 各クラスのパスを取得して読み込む
$iniClassPath = $pc.GetScriptPath("Ini_Class.ps1")
$pcClassPath = $pc.GetScriptPath("PC_Class.ps1")
$wordClassPath = $pc.GetScriptPath("Word_Class.ps1")

# IniFileクラスを読み込む
. $iniClassPath

# PCクラスを読み込む
. $pcClassPath

# Wordクラスを読み込む
. $wordClassPath

# メインクラスのインスタンスを作成して実行
$main = [Main]::new()