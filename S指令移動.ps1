# エラーハンドリングフラグの設定
# 0: エラーでも止めない
# 1: エラー都度止める
$errorHandlingFlag = 1

# デバッグモードの設定
$debugMode = $false

# トップフォルダパスの設定
$topFolderPath = "S:\技術部共有フォルダ\★一時交換フォルダ★\S指令"

# フォルダパスの設定
$sourceFolderPath = Join-Path -Path $topFolderPath -ChildPath "ExpFile"
$destinationTopFolderPath = "S:\技術部共有フォルダ\手配済みS指令[履歴]"
$copyListFilePath = Join-Path -Path $topFolderPath -ChildPath "CopyFile20241001115247.csv"
$logFolderPath = $topFolderPath

# ログファイルのパス
$logFilePath = Join-Path -Path $logFolderPath -ChildPath "copy_log.txt"
$errorLogFilePath = Join-Path -Path $logFolderPath -ChildPath "error_log.txt"

# ログファイルをクリア
if (Test-Path -Path $logFilePath) {
    Remove-Item -Path $logFilePath -Force
}
if (Test-Path -Path $errorLogFilePath) {
    Remove-Item -Path $errorLogFilePath -Force
}

# ログファイルにメッセージを追加する関数
function Write-Log {
    param (
        [string]$message
    )
    Add-Content -Path $logFilePath -Value $message
    if ($debugMode) {
        Write-Host $message
    }
}

# エラーログファイルにメッセージを追加する関数
function Write-ErrorLog {
    param (
        [string]$message
    )
    Add-Content -Path $errorLogFilePath -Value $message
    if ($debugMode) {
        Write-Host $message
    }
}

# CSVファイルの存在を確認
if (-not (Test-Path -Path $copyListFilePath)) {
    Write-ErrorLog "CSV file not found: $copyListFilePath"
    exit
}

# コピーリストを読み込む
$copyList = Import-Csv -Path $copyListFilePath -Encoding "shift_jis"

# コピー先フォルダを決定する関数
function Get-DestinationFolder {
    param (
        [string]$fileName
    )
    if ($fileName -match '^SS\d{6}$') {
        return "SS000000～"
    } else {
        $number = [int]($fileName -replace '\D', '')
        $baseNumber = [math]::Floor($number / 2000) * 2000
        return "S{0:000000}～" -f $baseNumber
    }
}

# 再帰的にファイルを検索する関数
function Find-File {
    param (
        [string]$folderPath,
        [string]$fileName
    )
    $file = Get-ChildItem -Path $folderPath -Recurse -Filter $fileName -ErrorAction SilentlyContinue
    return $file.FullName
}

# コピー先フォルダ毎のファイル数を集計するハッシュテーブル
$folderFileCount = @{}

# コピー処理
foreach ($row in $copyList) {
    $sourceFileName = $row.'ファイル名'
    $destinationFileName = $row.'タイトル'
    $sourceFilePath = Find-File -folderPath $sourceFolderPath -fileName $sourceFileName
    $destinationFolderName = Get-DestinationFolder -fileName $destinationFileName
    $destinationFolderPath = Join-Path -Path $destinationTopFolderPath -ChildPath $destinationFolderName
    $destinationFilePath = Join-Path -Path $destinationFolderPath -ChildPath $destinationFileName

    # コピー先フォルダが存在しない場合は作成
    if (-not (Test-Path -Path $destinationFolderPath)) {
        New-Item -Path $destinationFolderPath -ItemType Directory -Force
    }

    if ($sourceFilePath) {
        try {
            Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force
            Write-Log "Copied: $sourceFilePath to $destinationFilePath"

            # コピー先フォルダのファイル数を更新
            if ($folderFileCount.ContainsKey($destinationFolderName)) {
                $folderFileCount[$destinationFolderName]++
            } else {
                $folderFileCount[$destinationFolderName] = 1
            }
        } catch {
            Write-ErrorLog "Error copying $sourceFilePath to ${destinationFilePath}: $($_.Exception.Message)"
            if ($errorHandlingFlag -eq 1) {
                break
            }
        }
    } else {
        Write-ErrorLog "Source file not found: $sourceFileName"
        if ($errorHandlingFlag -eq 1) {
            break
        }
    }

    # 1秒待機
    Start-Sleep -Seconds 1

    # デバッグモードの場合、1ループ毎に停止
    if ($debugMode) {
        Read-Host "Press Enter to continue..."
    }
}

# コピー先フォルダ毎のファイル数をログ出力
foreach ($folder in $folderFileCount.Keys) {
    Write-Log "Folder: $folder, Files copied: $($folderFileCount[$folder])"
}