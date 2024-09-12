param (
    [string]$mode = "default",         # モード: "csv_only", "limited", "all", "specific"
    [int]$limit = 10,                  # 変換するファイルの数（"limited"モードの場合）
    [string]$specificFolder = "",      # 特定の変換元フォルダ（"specific"モードの場合）
    [switch]$debugMode                 # デバッグモードの有効化
)

# ネットワークドライブの設定
$networkPath = $null
$localDrive = "Z:"
# Remove the unused variable

if ($debugMode) {
    $networkPath = "\\DELLD033\技術部"
    $conversionFolder = "D:\eValue_TIFF2PDF_Debug\conversion"
    Write-Host "Debug mode enabled."
} else {
    $networkPath = "\\ycsvm112\技術部"
    $conversionFolder = "D:\eValue_TIFF2PDF\conversion"
    Write-Host "Production mode enabled."
}

# ネットワークドライブを一時的にマッピング
if (-not (Test-Path -Path $localDrive)) {
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $networkPath -Persist
}
$commonPath = "Z:\管理課\管理課共有資料\ArcSuite\eValue図面検索データ_240310\"

# 変換元フォルダのリスト
$sourceFolders = @(
    "図面検索【最新版】\図面\ExpFile",
    "図面検索【最新版】\通知書\ExpFile",
    "図面検索【最新版】\個装\ExpFile",
    "図面検索【旧版】\図面(旧)\ExpFile",
    "図面検索【旧版】\個装\ExpFile"
)

# デバッグ用の設定
if ($debugMode) {
    $debugTiffFiles = Get-ChildItem -Path $commonPath -Filter *.tiff | Select-Object -First 10
    $debugSourceFolders = $sourceFolders | Select-Object -First 3

    # TIFFファイルのコピー
    $debugTiffFiles | ForEach-Object {
        $destinationPath = Join-Path -Path $commonPath -ChildPath "DebugTiffFiles" -Resolve
        Copy-Item -Path $_.FullName -Destination $destinationPath
    }

    # デバッグ用の設定を使用する
    $commonPath = Join-Path -Path $commonPath -ChildPath "DebugTiffFiles"
    $sourceFolders = $debugSourceFolders
}

# CSVファイルの最大サイズ（MB）
$maxCsvFileSizeMB = 1

# PSCAN関連のフォルダパス
$pscanInFolder = Join-Path -Path $commonPath -ChildPath "PSCANの入力フォルダ"
$pscanOutFolder = Join-Path -Path $commonPath -ChildPath "PSCANの出力フォルダ"
$pscanOutErrFolder = Join-Path -Path $commonPath -ChildPath "PSCANのエラーフォルダ"

# OCR関連のフォルダパス
$OCR_OKFolder = Join-Path -Path $commonPath -ChildPath "OCR_OK"     # OCR成功フォルダ
$OCR_NGFolder = Join-Path -Path $commonPath -ChildPath "OCR_NG"     # OCR失敗フォルダ

# 変換リストの生成
function New-ConversionLists {
    param (
        [string]$sourceFolder
    )
    $conversionList = @()
    $tiffFiles = Get-ChildItem -Path $sourceFolder -Filter *.tiff
    $index = 1
    foreach ($tiffFile in $tiffFiles) {
        $conversionList += [PSCustomObject]@{
            Index = $index
            TIFFFile = $tiffFile.Name
            StartTime = ""
            Duration = ""
            Status = "Pending"
        }
        $index++ 
    }
    return $conversionList
}





# 変換リストの読み込み
function Get-ConversionList {
    param (
        [string]$sourceFolder
    )
    $csvFiles = Get-ChildItem -Path $sourceFolder -Filter "*_PDF変換リスト_*.csv" | Sort-Object Name
    $conversionList = @()
    foreach ($csvFile in $csvFiles) {
        $conversionList += Import-Csv -Path $csvFile
    }
    return $conversionList
}

# 変換リストの保存
function Save-ConversionList {
    param (
        [string]$sourceFolder,
        [array]$conversionList,
        [int]$fileIndex
    )
    $csvPath = Join-Path -Path $sourceFolder -ChildPath "$([System.IO.Path]::GetFileName($sourceFolder))_PDF変換リスト_$fileIndex.csv"
    $conversionList | Export-Csv -Path $csvPath -NoTypeInformation
}

# CSVファイルのサイズをチェックして新しいファイルを作成
function Test-CsvFileSize {
    param (
        [string]$csvPath,
        [array]$conversionList,
        [int]$fileIndex,
        [string]$sourceFolder
    )
    $csvFileSizeMB = (Get-Item $csvPath).Length / 1MB
    if ($csvFileSizeMB -ge $maxCsvFileSizeMB) {
        $fileIndex++
        $csvPath = Join-Path -Path $sourceFolder -ChildPath "$([System.IO.Path]::GetFileName($sourceFolder))_PDF変換リスト_$fileIndex.csv"
        $conversionList | Export-Csv -Path $csvPath -NoTypeInformation
    }
    return $csvPath, $fileIndex
}

# OCR結果の検証
function Test-OCR {
    param (
        [string]$pdfPath
    )
    $pdfText = (Get-Content $pdfPath -Raw)
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $numberPattern = "\d{6}"
    [regex]::Matches($fileName, $numberPattern)
    foreach ($match in $matches) {
        if ($pdfText -match $match.Value) {
            return $true
        }
    }
    return $false
}

# メイン処理
function Convert-Folder {
    param (
        [string]$sourceFolder,
        [string]$mode,
        [int]$limit
    )
    $destinationFolder = "$sourceFolder`_PDF"
    if (-not (Test-Path -Path $destinationFolder)) {
        New-Item -Path $destinationFolder -ItemType Directory
    }

    $conversionList = Get-ConversionList -sourceFolder $sourceFolder
    if (-not $conversionList) {
        $conversionList = New-ConversionList -sourceFolder $sourceFolder
    }

    $fileIndex = 1
    $csvPath = Join-Path -Path $sourceFolder -ChildPath "$([System.IO.Path]::GetFileName($sourceFolder))_PDF変換リスト_$fileIndex.csv"
    $conversionList | Export-Csv -Path $csvPath -NoTypeInformation

    $processedCount = 0

    foreach ($item in $conversionList) {
        if ($item.Status -eq "Pending") {
            $tiffFile = $item.TIFFFile
            $tiffPath = Join-Path -Path $sourceFolder -ChildPath $tiffFile
            $pdfFile = [System.IO.Path]::ChangeExtension($tiffFile, ".pdf")
            $pdfPath = Join-Path -Path $pscanOutFolder -ChildPath $pdfFile
            $destinationPath = Join-Path -Path $destinationFolder -ChildPath $pdfFile

            if (Test-Path -Path $tiffPath) {
                try {
                    $startTime = Get-Date
                    if ($mode -ne "csv_only") {
                        # TIFFファイルをPSCANの入力フォルダにコピー
                        Copy-Item -Path $tiffPath -Destination $pscanInFolder
                        # OCR処理が完了するまで待機
                        while (-not (Test-Path -Path $pdfPath) -and (Test-Path -Path $tiffPath)) {
                            Start-Sleep -Seconds 1
                        }
                        if (Test-Path -Path $pdfPath) {
                            # PDFファイルを完成フォルダに移動
                            Move-Item -Path $pdfPath -Destination $destinationPath
                            # OCR結果を検証
                            if (Test-OCR -pdfPath $destinationPath) {
                                Move-Item -Path $destinationPath -Destination $OCR_OKFolder
                                $item.Status = "OK"
                            } else {
                                Move-Item -Path $destinationPath -Destination $OCR_NGFolder
                                $item.Status = "NG"
                            }
                        } else {
                            # 変換が完了しなかった場合、エラーフォルダに移動
                            Move-Item -Path $tiffPath -Destination $pscanOutErrFolder
                            $item.Status = "NG"
                        }
                    }
                    $endTime = Get-Date
                    $duration = $endTime - $startTime

                    # 変換リストに成功を記録
                    $item.StartTime = $startTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $item.Duration = $duration.TotalSeconds
                }
                catch {
                    # 変換リストにエラーを記録
                    $item.Status = "NG"
                    break
                }
            } else {
                # 変換が完了しなかった場合、変換リストに失敗を記録
                $item.Status = "NG"
                break
            }

            $processedCount++
            if ($mode -eq "limited" -and $processedCount -ge $limit) {
                break
            }
        }
    }

    # 変換リストを保存
    $csvPath, $fileIndex = Test-CsvFileSize -csvPath $csvPath -conversionList $conversionList -fileIndex $fileIndex -sourceFolder $sourceFolder
    $conversionList | Export-Csv -Path $csvPath -NoTypeInformation
}

# 変換リストの生成
Generate-ConversionLists


# 処理の中断と再開
Write-Host "Processing stopped. Press any key to resume."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# フォルダの処理
$mode = "default"  # 必要に応じて変更
$limit = 100       # 必要に応じて変更
Process-Folders -mode $mode -limit $limit

# ネットワークドライブを解除
if (Get-PSDrive -Name "Z" -ErrorAction SilentlyContinue) {
    Remove-PSDrive -Name "Z" -Force
    Write-Host "Network drive Z: has been removed."
}