param (
    [switch]$debug,
    [switch]$SkipConverted = $true,
    [switch]$Reset,
    [switch]$ShowDialog
)

# スクリプトのディレクトリをホームディレクトリに設定
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $scriptDir


if ($ShowDialog) {
    Add-Type -AssemblyName System.Windows.Forms

    # 処理状況を表示するためのフォームを作成
    $form = New-Object Windows.Forms.Form
    $form.Text = "Processing Status"
    $form.Size = New-Object Drawing.Size(600, 400)
    $form.StartPosition = "CenterScreen"

    # DataGridViewを作成
    $dataGridView = New-Object Windows.Forms.DataGridView
    $dataGridView.Size = New-Object Drawing.Size(580, 350)
    $dataGridView.Location = New-Object Drawing.Point(10, 10)
    $dataGridView.ColumnCount = 3
    $dataGridView.Columns[0].Name = "FileName"
    $dataGridView.Columns[1].Name = "Status"
    $dataGridView.Columns[2].Name = "Timestamp"
    $form.Controls.Add($dataGridView)

    # フォームを表示
    $form.Show()

    # 処理状況を更新する関数
    function Update-GridView {
        param (
            [string]$fileName,
            [string]$status,
            [string]$timestamp
        )
        $dataGridView.Rows.Add($fileName, $status, $timestamp)
        $form.Refresh()
    }
} else {
    # ダミーの関数を定義しておく
    function Update-GridView {
        param (
            [string]$fileName,
            [string]$status,
            [string]$timestamp
        )
    }
}

# ネットワークドライブの設定
$networkPath = if ($debug) { "\\DELLD033\技術部" } else { "\\ycsvm112\技術部" }
$localDrive = "Z:"
$commonPath = "Z:\管理課\管理課共有資料\ArcSuite\eValue図面検索データ_240310\"

# ネットワークドライブを一時的にマッピング
if (-not (Test-Path -Path $localDrive)) {
    Write-Host "Mapping network drive..."
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $networkPath -Persist
}

# ネットワークドライブの確認
if (-not (Test-Path -Path $commonPath)) {
    Write-Host "Error: Network drive not mapped correctly or path does not exist."
    exit
}

# 変換元フォルダのリスト（全角の「￥」をそのまま使用）
$sourceFolders = @(
    "図面検索【最新版】￥図面",
    "図面検索【最新版】￥通知書",
    "図面検索【最新版】￥個装",
    "図面検索【旧版】￥図面(旧)",
    "図面検索【旧版】￥個装"
)

# デバッグモードの設定
$maxFilesToProcess = if ($debug) { 1 } else { [int]::MaxValue }

# フォルダのパス設定
$pscanInFolder = "\\10.23.2.28\HGPscanServPlus5\Job02_OCR\OCR_IN"
$pscanOutFolder = "\\10.23.2.28\HGPscanServPlus5\Job02_OCR\OCR_OUT"

# 変換リストの生成
$tiffListPath = Join-Path -Path $commonPath -ChildPath "TIFF_LIST.txt"
if ($Reset -and (Test-Path -Path $tiffListPath)) {
    Remove-Item -Path $tiffListPath -Force
}
if (-not (Test-Path -Path $tiffListPath)) {
    New-Item -Path $tiffListPath -ItemType File
}

# 待機時間の設定
$timeoutMinutes = 1
$sleepSeconds = 10

# リセットモードの場合、PDFフォルダとERRフォルダを削除
if ($Reset) {
    foreach ($sourceFolder in $sourceFolders) {
        $sourceFolderPath = Join-Path $commonPath -ChildPath $sourceFolder
        $pdfFolderBase = Join-Path -Path $sourceFolderPath -ChildPath "PDF"
        $pdfErrFolderBase = Join-Path -Path $sourceFolderPath -ChildPath "ERR"
        
        if (Test-Path -Path $pdfFolderBase) {
            Remove-Item -Path $pdfFolderBase -Recurse -Force
        }
        if (Test-Path -Path $pdfErrFolderBase) {
            Remove-Item -Path $pdfErrFolderBase -Recurse -Force
        }
    }
}

# TIFFファイルの抽出と処理
foreach ($sourceFolder in $sourceFolders) {
    $sourceFolderPath = Join-Path $commonPath -ChildPath $sourceFolder
    $tiffFolder = Join-Path -Path $sourceFolderPath -ChildPath "ExpFile"
    $pdfFolderBase = Join-Path -Path $sourceFolderPath -ChildPath "PDF"
    $pdfErrFolderBase = Join-Path -Path $sourceFolderPath -ChildPath "ERR"
        
    if (-not (Test-Path -Path $tiffFolder)) {
        Write-Host "Error: Path does not exist - $tiffFolder"
        continue
    }

    $tiffFiles = Get-ChildItem -Path $tiffFolder -Filter *.tif | Select-Object -First $maxFilesToProcess

    foreach ($tiffFile in $tiffFiles) {
        $tiffFileName = $tiffFile.Name
        $tiffFilePath = $tiffFile.FullName
        $pdfFileName = [System.IO.Path]::ChangeExtension($tiffFileName, ".pdf")
        $pdfFilePath = Join-Path -Path $pdfFolderBase -ChildPath $pdfFileName

        # 既に変換済みのTIFFファイルをスキップ
        if ($SkipConverted -and (Test-Path -Path $pdfFilePath)) {
            Write-Host "Skipping already converted file: $tiffFileName"
            continue
        }

        # TIFFファイルをPSCAN_INにコピー
        Copy-Item -Path $tiffFilePath -Destination $pscanInFolder

        # TIFF_LISTに追記
        Add-Content -Path $tiffListPath -Value "$sourceFolder`t$tiffFileName`t" -NoNewline

        # PDFファイルが生成されるかを1分間（10秒ごとにチェック）待機する処理
        $pdfFilePath = Join-Path -Path $pscanOutFolder -ChildPath $pdfFileName
        $waittime = [datetime]::Now.AddMinutes($timeoutMinutes)
        while ((-not (Test-Path -Path $pdfFilePath)) -and ([datetime]::Now -lt $waittime)) {
            Start-Sleep -Seconds $sleepSeconds
        }

        if (Test-Path -Path $pdfFilePath) {
            # PDFフォルダの作成
            if (-not (Test-Path -Path $pdfFolderBase)) {
                New-Item -Path $pdfFolderBase -ItemType Directory
            }

            # PDFファイルを移動（既存のファイルがある場合は上書き）
            Move-Item -Path $pdfFilePath -Destination $pdfFolderBase -Force

            # TIFF_LISTにOKを追記
            $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Add-Content -Path $tiffListPath -Value "OK`t$timestamp`n"

            # 処理状況を更新
            Update-GridView -fileName $tiffFileName -status "OK" -timestamp $timestamp
        } else {
            # タイムアウト処理
            if (-not (Test-Path -Path $pdfErrFolderBase)) {
                New-Item -Path $pdfErrFolderBase -ItemType Directory
            }

            # TIFFファイルをERRフォルダに移動（既存のファイルがある場合は上書き）
            Move-Item -Path $tiffFilePath -Destination $pdfErrFolderBase -Force

            # TIFF_LISTにNGを追記
            $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Add-Content -Path $tiffListPath -Value "NG`t$timestamp`n"

            # 処理状況を更新
            Update-GridView -fileName $tiffFileName -status "NG" -timestamp $timestamp
        }
    }
}

# ネットワークドライブを解除
if (Get-PSDrive -Name "Z" -ErrorAction SilentlyContinue) {
    Remove-PSDrive -Name "Z" -Force
    Write-Host "Network drive Z: has been removed."
}