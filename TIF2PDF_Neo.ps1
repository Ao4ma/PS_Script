# ログファイルとエラーファイルのパスを設定
$logFilePath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\UpPDF\log.txt"
$errorFilePath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\UpPDF\error.txt"

# ログファイルとエラーファイルのディレクトリが存在するか確認し、存在しない場合は作成
$destinationBasePath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\UpPDF\生成PDF"
if (-not (Test-Path -Path $destinationBasePath)) {
    New-Item -Path $destinationBasePath -ItemType Directory
}

# TIFファイルを再帰的に検索する関数
function Find-TifFiles {
    param (
        [string]$topFolderPath
    )

    # TIFファイルのパスを再帰的に検索
    $tifFilePaths = Get-ChildItem -Path $topFolderPath -Recurse -Include *.tif, *.tiff -ErrorAction SilentlyContinue
    return $tifFilePaths
}

# PDFファイルを検索してコピーし、TIFファイルを移動する関数
function Process-TifFiles {
    param (
        [string]$tifFilePath, # TIFファイルのパス
        [string]$pdfBasePath, # PDFファイルのベースパス
        [string]$logFilePath, # ログファイルのパス
        [string]$errorFilePath # エラーファイルのパス
    )

    $tifFileName = [System.IO.Path]::GetFileNameWithoutExtension($tifFilePath)
    $pdfFilePath = Get-ChildItem -Path $pdfBasePath -Recurse -Include "$tifFileName.pdf" -ErrorAction SilentlyContinue

    if ($pdfFilePath) {
        $tifFolderPath = [System.IO.Path]::GetDirectoryName($tifFilePath)
        $destinationPdfPath = Join-Path -Path $tifFolderPath -ChildPath "$tifFileName.pdf"
        Copy-Item -Path $pdfFilePath.FullName -Destination $destinationPdfPath -Force

        # 消去可能フォルダの作成
        $deletableFolderName = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName([System.IO.Path]::GetDirectoryName($tifFolderPath)))
        $deletableFolderPath = Join-Path -Path $tifFolderPath -ChildPath "..\..\消去可能_$deletableFolderName"
        if (-not (Test-Path -Path $deletableFolderPath)) {
            New-Item -Path $deletableFolderPath -ItemType Directory
        }

        # TIFファイルを消去可能フォルダに移動
        $destinationTifPath = Join-Path -Path $deletableFolderPath -ChildPath ([System.IO.Path]::GetFileName($tifFilePath))
        Move-Item -Path $tifFilePath -Destination $destinationTifPath -Force

        # ログに記録
        Add-Content -Path $logFilePath -Value "Copied PDF: $pdfFilePath to $destinationPdfPath and moved TIF: $tifFilePath to $destinationTifPath"
    } else {
        # PDFが見つからなかった場合、エラーログに記録
        Add-Content -Path $errorFilePath -Value "PDF not found for TIF: $tifFilePath"
    }
}

# メイン処理
$topFolderPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\UpPDF\#eValue元データtiff為し"
$pdfBasePath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\UpPDF\生成PDF"

# TIFファイルを再帰的に検索
$tifFiles = Find-TifFiles -topFolderPath $topFolderPath

# 各TIFファイルを処理
foreach ($tifFile in $tifFiles) {
    Process-TifFiles -tifFilePath $tifFile.FullName -pdfBasePath $pdfBasePath -logFilePath $logFilePath -errorFilePath $errorFilePath
}