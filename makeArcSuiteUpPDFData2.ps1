# FileManagerクラスの定義
class FileManager {
    # CSVに基づいてファイルをコピーするメソッド
    [void] CopyFilesBasedOnCsv([string]$csvFilePath, [string]$workHomeFolder, [string]$outDataFolder, [string]$realDataFolder, [ref]$successCount, [ref]$failureCount) {
        # CSVファイルをインポート
        $csvData = Import-Csv -Path $csvFilePath

        foreach ($record in $csvData) {
            # CSVの1列目をファイル名として取得
            $fileName = $record.PSObject.Properties[0].Value

            # realDataFolderを再帰的にサーチしてファイルを探す
            $foundFile = Get-ChildItem -Path $realDataFolder -Recurse -Filter $fileName | Select-Object -First 1

            if ($foundFile) {
                # outDataFolder直下にCSVファイル名のフォルダを作成
                $csvFolderName = [System.IO.Path]::GetFileNameWithoutExtension($csvFilePath)
                $destinationFolder = Join-Path -Path $outDataFolder -ChildPath $csvFolderName

                if (-Not (Test-Path -Path $destinationFolder)) {
                    New-Item -Path $destinationFolder -ItemType Directory
                }

                # 見つけたファイルをコピー
                $destinationPath = Join-Path -Path $destinationFolder -ChildPath $foundFile.Name
                Copy-Item -Path $foundFile.FullName -Destination $destinationPath -Force
                $successCount.Value++
                $this.LogSuccess($workHomeFolder, "Successfully copied: $fileName to $destinationPath")
                Write-Host "Successfully copied: $fileName to $destinationPath"
            } else {
                # ファイルが見つからなかった場合の処理
                $this.LogError($workHomeFolder, "File not found: $fileName, Record: $($record | ConvertTo-Json -Compress)")
                $failureCount.Value++
                Write-Host "File not found: $fileName"
            }
        }
    }

    # フォルダが存在するか確認し、存在しない場合は作成するメソッド
    [void] EnsureFolderExists([string]$folderPath) {
        if (-Not (Test-Path -Path $folderPath)) {
            New-Item -Path $folderPath -ItemType Directory
        } else {
            Remove-Item -Path "$folderPath\*" -Recurse -Force
        }
    }

    # エラーログを記録するメソッド
    [void] LogError([string]$workHomeFolder, [string]$message) {
        $errorFile = Join-Path -Path $workHomeFolder -ChildPath "error.log"
        Add-Content -Path $errorFile -Value $message
    }

    # 成功ログを記録するメソッド
    [void] LogSuccess([string]$workHomeFolder, [string]$message) {
        $successFile = Join-Path -Path $workHomeFolder -ChildPath "success.log"
        Add-Content -Path $successFile -Value $message
    }
}

# PCクラスの定義
class PC {
    [string]$Name
    [string]$IPAddress
    [string]$HomeFolder
    [string]$WorkFolder

    # コンストラクタ
    PC([string]$name, [string]$ipAddress) {
        $this.Name = $name
        $this.IPAddress = $ipAddress

        # コマンドプロンプトでhostnameを実行してPC名を取得
        $pcName = (hostname).Trim()

        # 特定のPC名を別名に変換
        switch ($pcName) {
            "TUF-FX517ZM" { $pcName = "AsusTUF" }
        }

        # PC名に応じたホームフォルダとワークフォルダの設定
        switch ($pcName) {
            "Delld033" {
                $this.HomeFolder = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\"
                $this.WorkFolder = "S:\\技術部storage\\管理課\\PDM復旧"
            }
            "AsusTUF" {
                $this.HomeFolder = "C:\\Users\\22200\\OneDrive\\ドキュメント\\GitHub\\PS_Script\\"
                $this.WorkFolder = "D:\\技術部storage\\管理課\\PDM復旧"
            }
            default {
                throw "Unknown PC name: $pcName"
            }
        }
    }

    # オブジェクトの文字列表現を返すメソッド
    [string] ToString() {
        return "$($this.Name) ($($this.IPAddress))"
    }
}

# エスケープ文字処理の例
function EscapeSpecialCharacters {
    param (
        [string]$inputString
    )
    return $inputString -replace '([\\*+?.()|{}[\]])', '\\$1'
}

# メイン処理
$fileManager = [FileManager]::new()
$successCount = [ref]0
$failureCount = [ref]0

# ファイルコピー処理の実行
$fileManager.CopyFilesBasedOnCsv("C:\path\to\your\csvfile.csv", $settings.ScriptHomeFolder, $settings.OutDataFolder, "C:\RealDataFolder", [ref]$successCount, [ref]$failureCount)

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"
