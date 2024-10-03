# デバッグモードの設定
$debugMode = $false

# トップフォルダパスの設定
$topFolderPath = "S:\技術部共有フォルダ\★一時交換フォルダ★\S指令"

# フォルダパスの設定
$destinationTopFolderPath = "S:\技術部共有フォルダ\手配済みS指令[履歴]"
$copyListFilePath = Join-Path -Path $topFolderPath -ChildPath "CopyFile20241001115247.csv"
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

# CSVファイルの存在を確認
if (-not (Test-Path -Path $copyListFilePath)) {
    Write-CheckErrorLog "CSV file not found: $copyListFilePath"
    exit
}

# コピーリストを読み込む
$copyList = Import-Csv -Path $copyListFilePath -Encoding "shift_jis"

# フォルダごとに比較照合
$folders = $copyList | Select-Object -ExpandProperty 'フォルダパス' | Sort-Object -Unique

foreach ($folderPath in $folders) {
    $folder = Split-Path -Path $folderPath -Leaf
    $selectedFolderPath = Join-Path -Path $destinationTopFolderPath -ChildPath $folder

    # ログファイルのパスを決定
    $logFilePath = Join-Path -Path $logFolderPath -ChildPath "copy_log_$folder.txt"

    # ログファイルの存在を確認
    if (-not (Test-Path -Path $logFilePath)) {
        Write-CheckErrorLog "Log file not found: $logFilePath"
        continue
    }

    # ログファイルを読み込む
    $logEntries = Get-Content -Path $logFilePath

    # フォルダ内のファイル数を取得
    $actualFileCount = (Get-ChildItem -Path $selectedFolderPath -File -Recurse).Count

    # CSVファイルのsourceFileNameの数を取得
    $csvFileCount = ($copyList | Where-Object { $_.'フォルダパス' -eq $folderPath }).Count

    # ログファイルの実施ログを一行ずつ確認
    foreach ($logEntry in $logEntries) {
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
        $destinationFolder = Split-Path -Path $destinationFilePath -Parent
        if ($destinationFolder -ne $selectedFolderPath) {
            Write-CheckErrorLog "Incorrect destination folder: $destinationFolder (expected: $selectedFolderPath)"
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

    # ログファイルの行数を取得
    $logFileLineCount = $logEntries.Count

    # 比較結果を照合ログに出力
    Write-CheckLog -folderPath $selectedFolderPath -message "Comparing log file: $logFilePath with folder: $selectedFolderPath"
    Write-CheckLog -folderPath $selectedFolderPath -message "Actual files in folder: $actualFileCount"
    Write-CheckLog -folderPath $selectedFolderPath -message "CSV source file count: $csvFileCount"
    Write-CheckLog -folderPath $selectedFolderPath -message "Log file line count: $logFileLineCount"

    # ターミナル出力
    Write-Host "Folder: $folder`tActual: $actualFileCount`tCSV: $csvFileCount`tLog lines: $logFileLineCount"

    # ログファイルの行数とコピー先のファイル数が一致するか確認
    if ($actualFileCount -ne $logFileLineCount) {
        Write-CheckErrorLog "Mismatch between actual file count ($actualFileCount) and log file line count ($logFileLineCount)"
    }
}

# CSVリストの内容がすべてコピーされているか確認
$totalCsvFiles = $copyList.Count
$totalActualFiles = (Get-ChildItem -Path $destinationTopFolderPath -File -Recurse).Count

Write-CheckLog -folderPath $destinationTopFolderPath -message "Total files specified in CSV: $totalCsvFiles"
Write-CheckLog -folderPath $destinationTopFolderPath -message "Total actual files in all folders: $totalActualFiles"

# ターミナル出力
Write-Host "Total files specified in CSV: $totalCsvFiles"
Write-Host "Total actual files in all folders: $totalActualFiles"

# CSVリストの内容がすべてコピーされているか確認
if ($totalCsvFiles -ne $totalActualFiles) {
    Write-CheckErrorLog "Mismatch between total files specified in CSV ($totalCsvFiles) and total actual files in all folders ($totalActualFiles)"
}