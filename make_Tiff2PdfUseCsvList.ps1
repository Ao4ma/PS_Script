param (
    [string]$csvPath = "C:\ps1\_新旧リスト比較_pdf未変換.csv", # CSVファイルのパス
    [string]$topFolderPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)", # トップフォルダのパス
    [string]$destinationBasePath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdfデータ\追加生成PDF" # 目的地のベースパス
)

# ログファイルとエラーファイルのパスを設定
$logFilePath = Join-Path -Path $destinationBasePath -ChildPath "log.txt" # ログファイルのパス
$errorFilePath = Join-Path -Path $destinationBasePath -ChildPath "error.txt" # エラーファイルのパス

# ログファイルとエラーファイルのディレクトリが存在するか確認し、存在しない場合は作成
if (-not (Test-Path -Path $destinationBasePath)) {
    New-Item -Path $destinationBasePath -ItemType Directory
}

function Find-ExpFilePaths {
    param (
        [string]$topFolderPath
    )

    # ExpFileフォルダのパスを再帰的に検索
    $expFilePaths = Get-ChildItem -Path $topFolderPath -Recurse -Directory -Filter "ExpFile" -ErrorAction SilentlyContinue
    return $expFilePaths
}

function Copy-FilesIfExist {
    param (
        [string]$csvPath, # CSVファイルのパス
        [string]$topFolderPath, # トップフォルダのパス
        [string]$destinationBasePath, # 目的地のベースパス
        [string]$logFilePath, # ログファイルのパス
        [string]$errorFilePath # エラーファイルのパス
    )

    # CSVデータをインポート（エンコーディングを指定）
    $csvData = Import-Csv -Path $csvPath -Encoding UTF8

    # ExpFileフォルダのパスを取得
    $expFilePaths = Find-ExpFilePaths -topFolderPath $topFolderPath

    # ExpFileフォルダのパスを表示
    foreach ($expFilePath in $expFilePaths) {
        Write-Output "Found ExpFile folder: $($expFilePath.FullName)"
    }

    # 各行に対して処理を実行
    foreach ($row in $csvData) {
        $folderPath = $row.フォルダパス # フォルダパスを取得
        $extension = $row.拡張子 # 拡張子を取得
        $oldFileName = $row.ファイル名_変更前 # 変更前のファイル名を取得

        # ExpFileフォルダ内でファイルを検索
        $fileFound = $false
        foreach ($expFilePath in $expFilePaths) {
            if (Test-Path -Path $expFilePath.FullName) {
                $filePath = Get-ChildItem -Path $expFilePath.FullName -Recurse -Filter "$oldFileName.$extension" -ErrorAction SilentlyContinue
                if ($filePath) {
                    $fileFound = $true
                    $destinationFolder = Join-Path -Path $destinationBasePath -ChildPath $folderPath.TrimStart("\")
                    if (-not (Test-Path -Path $destinationFolder)) {
                        New-Item -Path $destinationFolder -ItemType Directory # 目的地フォルダが存在しない場合は作成
                    }
                    try {
                        Copy-Item -Path $filePath.FullName -Destination $destinationFolder -ErrorAction Stop # ファイルをコピー
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $logMessage = "[$timestamp] Copied: $filePath to $destinationFolder"
                        Write-Output $logMessage
                        Add-Content -Path $logFilePath -Value $logMessage # ログに記録
                    } catch {
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $errorMessage = "[$timestamp] Failed to copy: $filePath to $destinationFolder"
                        Write-Output $errorMessage
                        # エラーファイルのディレクトリが存在するか確認し、存在しない場合は作成
                        if (-not (Test-Path -Path (Split-Path -Path $errorFilePath -Parent))) {
                            New-Item -Path (Split-Path -Path $errorFilePath -Parent) -ItemType Directory
                        }
                        Add-Content -Path $errorFilePath -Value $errorMessage # エラーを記録
                    }
                    break
                }
            }
        }

        if (-not $fileFound) {
            # ExpFileフォルダ内で見つからなかった場合、トップフォルダから再帰的に検索
            $filePath = Get-ChildItem -Path $topFolderPath -Recurse -Filter "$oldFileName.$extension" -ErrorAction SilentlyContinue
            if ($filePath) {
                $destinationFolder = Join-Path -Path $destinationBasePath -ChildPath $folderPath.TrimStart("\")
                if (-not (Test-Path -Path $destinationFolder)) {
                    New-Item -Path $destinationFolder -ItemType Directory # 目的地フォルダが存在しない場合は作成
                }
                try {
                    Copy-Item -Path $filePath.FullName -Destination $destinationFolder -ErrorAction Stop # ファイルをコピー
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $logMessage = "[$timestamp] Copied: $filePath to $destinationFolder"
                    Write-Output $logMessage
                    Add-Content -Path $logFilePath -Value $logMessage # ログに記録
                } catch {
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $errorMessage = "[$timestamp] Failed to copy: $filePath to $destinationFolder"
                    Write-Output $errorMessage
                    # エラーファイルのディレクトリが存在するか確認し、存在しない場合は作成
                    if (-not (Test-Path -Path (Split-Path -Path $errorFilePath -Parent))) {
                        New-Item -Path (Split-Path -Path $errorFilePath -Parent) -ItemType Directory
                    }
                    Add-Content -Path $errorFilePath -Value $errorMessage # エラーを記録
                }
            } else {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $errorMessage = "[$timestamp] File not found: $oldFileName.$extension in any ExpFile folders under $topFolderPath"
                Write-Output $errorMessage
                # エラーファイルのディレクトリが存在するか確認し、存在しない場合は作成
                if (-not (Test-Path -Path (Split-Path -Path $errorFilePath -Parent))) {
                    New-Item -Path (Split-Path -Path $errorFilePath -Parent) -ItemType Directory
                }
                Add-Content -Path $errorFilePath -Value $errorMessage # エラーを記録
            }
        }
    }
}

# 関数を実行
Copy-FilesIfExist -csvPath $csvPath -topFolderPath $topFolderPath -destinationBasePath $destinationBasePath -logFilePath $logFilePath -errorFilePath $errorFilePath