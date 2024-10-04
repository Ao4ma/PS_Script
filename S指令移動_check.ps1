# デバッグモードの設定
$debugMode = $true

# トップフォルダパスの設定
$topFolderPath = "S:\技術部共有フォルダ\★一時交換フォルダ★\S指令"

# フォルダパスの設定
$destinationTopFolderPath = "S:\技術部共有フォルダ\手配済みS指令`[履歴`]"

# ログフォルダパスの設定
$logFolderPath = Join-Path -Path $topFolderPath -ChildPath "ログ"
$checklogFolderPath = Join-Path -Path $topFolderPath -ChildPath "照合ログ"

# 照合ログフォルダを作成
if (-not (Test-Path -Path $checklogFolderPath)) {
    New-Item -Path $checklogFolderPath -ItemType Directory -Force
}

# 照合エラーログファイルのパス
$checkerrorLogFilePath = Join-Path -Path $checklogFolderPath -ChildPath "照合error_log.txt"

# 照合エラーログファイルをクリア
if (Test-Path -Path $checkerrorLogFilePath) {
    Remove-Item -Path $checkerrorLogFilePath -Force
}

# 照合ログファイルにメッセージを追加する関数
function Write-CheckLog {
    param (
        [string]$folderPath,
        [string]$message
    )
    $folderName = Split-Path -Path $folderPath -Leaf
    $checklogFileName = "照合copy_log_$folderName.txt"
    $checklogFilePath = Join-Path -Path $checklogFolderPath -ChildPath $checklogFileName
    Add-Content -Path $checklogFilePath -Value $message
    if ($debugMode) {
        Write-Host $message
    }
}

# 照合エラーログファイルにメッセージを追加する関数
function Write-CheckErrorLog {
    param (
        [string]$message
    )
    Add-Content -Path $checkerrorLogFilePath -Value $message
    if ($debugMode) {
        Write-Host $message
    }
}

# コピー先フォルダを取得
$destinationFolders = Get-ChildItem -Path "S:\技術部共有フォルダ\手配済みS指令``[履歴``]" -Directory

foreach ($destinationFolder in $destinationFolders) {
    $folder = $destinationFolder.Name

    # ログファイルを検索
    Write-Host "Searching for log file for folder: $folder"
    $logFilePath = Get-ChildItem -Path $logFolderPath -Filter "copy_log_$folder.txt"

    # ログファイルの存在を確認
    if (-not $logFilePath) {
        Write-CheckErrorLog "Log file not found for folder: $folder"
        continue
    }

    # ログファイルを読み込む
    Write-Host "Reading log file: $logFilePath"
    $logEntries = Get-Content -Path $logFilePath.FullName

    # フォルダ内のファイル数を取得
    Write-Host "Counting files in folder: $destinationFolder"
    $actualFileCount = (Get-ChildItem -Path "$($destinationFolder.FullName)" -File -Recurse).Count

    # デバッグ用にファイルリストを表示
    $files = Get-ChildItem -Path "$($destinationFolder.FullName)" -File -Recurse
    Write-Host "Files in folder $($destinationFolder.FullName):"
    foreach ($file in $files) {
        Write-Host $file.FullName
    }

    # ログファイルの実施ログを一行ずつ確認（最後の1行を除外）
    for ($i = 0; $i -lt $logEntries.Count - 1; $i++) {
        $logEntry = $logEntries[$i]
        $logParts = $logEntry -split '\s+'
        $sourceFilePath = $logParts[1]
        $destinationFilePath = $logParts[3]

        # コピー元のファイルが存在するか確認
        if (-not (Test-Path -Path $sourceFilePath)) {
            Write-CheckErrorLog "Source file not found: $sourceFilePath"
        }

        # コピー先のファイルが存在するか確認
        if (-not (Test-Path -Path $destinationFilePath)) {
            Write-CheckErrorLog "Destination file not found: $destinationFilePath"
        }

        # コピー先フォルダが正しいか確認
        if ($destinationFilePath) {
            $destinationFolderPath = Split-Path -Path $destinationFilePath -Parent
            if ($destinationFolderPath -ne $destinationFolder.FullName) {
                Write-CheckErrorLog "Incorrect destination folder: $destinationFolderPath (expected: $destinationFolder.FullName)"
            }
        }

        # ファイル名の数字部分がフォルダの範囲内にあるか確認
        $fileName = Split-Path -Path $destinationFilePath -Leaf
        if ($fileName -match '^\D*(\d+)\D*$') {
            $fileNumber = [int]$matches[1]
            if ($folder -match '^\D*(\d+)\D*$') {
                $folderNumber = [int]$matches[1]
                if ($fileName -like 'SS*') {
                    if ($folder -ne 'SS000000') {
                        Write-CheckErrorLog "File $fileName should be in SS000000 folder, but found in $folder"
                    }
                } else {
                    if ($fileNumber -lt $folderNumber -or $fileNumber -ge ($folderNumber + 2000)) {
                        Write-CheckErrorLog "File $fileName is out of range for folder $folder"
                    }
                }
            }
        }
    }

    # ログファイルの行数を取得（最後の1行を除外）
    $logFileLineCount = $logEntries.Count - 1

    # 比較結果を照合ログに出力
    Write-CheckLog -folderPath $destinationFolder.FullName -message "Comparing log file: $($logFilePath.FullName) with folder: $($destinationFolder.FullName)"
    Write-CheckLog -folderPath $destinationFolder.FullName -message "Actual files in folder: $actualFileCount"
    Write-CheckLog -folderPath $destinationFolder.FullName -message "Log file line count: $logFileLineCount"

    # ターミナル出力
    Write-Host "Folder: $folder`tActual: $actualFileCount`tLog lines: $logFileLineCount"

    # ログファイルの行数とコピー先のファイル数が一致するか確認
    if ($actualFileCount -ne $logFileLineCount) {
        Write-CheckErrorLog "Mismatch between actual file count ($actualFileCount) and log file line count ($logFileLineCount)"
    }
}