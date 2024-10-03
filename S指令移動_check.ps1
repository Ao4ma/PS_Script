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

# コピー先フォルダを1つ選択（例：S002000～）
$selectedFolder = "S002000～"
$selectedFolderPath = Join-Path -Path $destinationTopFolderPath -ChildPath $selectedFolder

# ログファイルのパスを決定
$logFilePath = Join-Path -Path $logFolderPath -ChildPath "copy_log_$selectedFolder.txt"

# ログファイルの存在を確認
if (-not (Test-Path -Path $logFilePath)) {
    Write-CheckErrorLog "Log file not found: $logFilePath"
    exit
}

# ログファイルを読み込む
$logEntries = Get-Content -Path $logFilePath

# フォルダ内のファイル数を取得
$actualFileCount = (Get-ChildItem -Path $selectedFolderPath -File -Recurse).Count

# CSVファイルのsourceFileNameの数を取得
$csvFileCount = ($copyList | Where-Object { $_.'フォルダ名' -eq $selectedFolder }).Count

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
}

# ログファイルの行数を取得
$logFileLineCount = $logEntries.Count

# 比較結果を照合ログに出力
Write-CheckLog -folderPath $selectedFolderPath -message "Actual files in folder: $actualFileCount"
Write-CheckLog -folderPath $selectedFolderPath -message "CSV source file count: $csvFileCount"
Write-CheckLog -folderPath $selectedFolderPath -message "Log file line count: $logFileLineCount"

# ターミナル出力
Write-Host "Folder: $selectedFolder`tActual: $actualFileCount`tCSV: $csvFileCount`tLog lines: $logFileLineCount"

# ログファイルの行数とコピー先のファイル数が一致するか確認
if ($actualFileCount -ne $logFileLineCount) {
    Write-CheckErrorLog "Mismatch between actual file count ($actualFileCount) and log file line count ($logFileLineCount)"
}