# モジュールをインポート
using module ./MyLibrary/PC_Class.psm1
using module ./MyLibrary/Ini_Class.psm1
using module ./MyLibrary/Word_Class.psm1

# スクリプトのフォルダパスを取得
$scriptFolderPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# スクリプトのフォルダパスを取得
$iniFilePath = Join-Path -Path $scriptFolderPath -ChildPath "config_Change_Properties_Word.ini"

class Main {
    [string]$scriptFolderPath
    [string]$iniFilePath

    Main([string]$scriptFolderPath, [string]$iniFilePath) {
        $this.scriptFolderPath = $scriptFolderPath
        $this.iniFilePath = $iniFilePath
    }

    [void]Run() {
        # MyPCクラスのインスタンスを作成
        $pc = [MyPC]::new((hostname))
        $pc.SetScriptFolder($this.scriptFolderPath)
        $pc.SetLogFolder("$this.scriptFolderPath\Logs")

        # PC情報を表示
        $pc.DisplayInfo()
        Write-Host "IP Address: $($pc.GetIPAddress())"
        Write-Host "MAC Address: $($pc.GetMACAddress())"
        Write-Host "Installed Libraries:"
        $pc.ListInstalledLibraries()

        # IniFileクラスのインスタンスを作成
        $ini = [IniFile]::new($this.iniFilePath)

        # PC情報を調べて、iniファイルに記録
        $this.RecordPCInfo($pc, $ini)

        # ドキュメントのパスを取得
        $docFilePath = $ini.GetValue("DocFile", "DocFilePath")
        $docFileName = $ini.GetValue("DocFile", "DocFileName")
        $filePath = Join-Path -Path $docFilePath -ChildPath $docFileName

        # ドキュメントを処理
        $this.ProcessDocument($pc, $filePath)
    }

    [void]RecordPCInfo([MyPC]$pc, [IniFile]$ini) {
        $pcInfo = @{
            "OS" = (Get-WmiObject -Class Win32_OperatingSystem).Caption
            "UserName" = $env:USERNAME
            "MachineName" = $env:COMPUTERNAME
        }

        foreach ($key in $pcInfo.Keys) {
            $ini.SetValue("PCInfo", $key, $pcInfo[$key])
        }
    }

    [void]ProcessDocument([MyPC]$pc, [string]$filePath) {
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
        }
    }
}

# モジュールをインポートした後にクラスを使用する
$main = [Main]::new($scriptFolderPath, $iniFilePath)
$main.Run()