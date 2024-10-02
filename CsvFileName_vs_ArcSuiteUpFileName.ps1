# エラーハンドリングフラグの設定
# 0: エラーでも止めない
# 1: エラー都度止める
# 2: CSVファイル単位で止める
$errorHandlingFlag = 0

# フォルダパスの設定
$csvFolderPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用csvデータ"
$pdfFolderPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdfデータ"
$errorLogFolderPath = Join-Path -Path $csvFolderPath -ChildPath "確認"
$stateFilePath = Join-Path -Path $errorLogFolderPath -ChildPath "state.json"

# エラーログフォルダが存在しない場合は作成
if (-not (Test-Path -Path $errorLogFolderPath)) {
    New-Item -Path $errorLogFolderPath -ItemType Directory -Force
} else {
    # エラーログフォルダ内の既存のログファイルを削除
    Get-ChildItem -Path $errorLogFolderPath -Filter *.log | Remove-Item -Force
}

# 中断地点の情報を読み込む
$state = @{}
if (Test-Path -Path $stateFilePath) {
    $state = Get-Content -Path $stateFilePath | ConvertFrom-Json
    $state = @{} + $state.PSObject.Properties | ForEach-Object { $_.Name = $state.$($_.Name) }
}

# CSVファイルを取得
$csvFiles = Get-ChildItem -Path $csvFolderPath -Filter *.csv

foreach ($csvFile in $csvFiles) {
    $csvFileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
    $pdfFolderForCsv = Join-Path -Path $pdfFolderPath -ChildPath $csvFileNameWithoutExtension

    # CSVファイル名と同名のフォルダが存在するか確認
    if (Test-Path -Path $pdfFolderForCsv) {
        $csvData = Import-Csv -Path $csvFile.FullName
        $errorMessages = @()

        # 中断地点から再開するためのインデックス
        $startIndex = if ($state.ContainsKey($csvFileNameWithoutExtension)) { $state[$csvFileNameWithoutExtension] } else { 0 }

        for ($i = $startIndex; $i -lt $csvData.Count; $i++) {
            $row = $csvData[$i]

            # '関連付け用ファイル名' 列が存在するか確認
            if (-not $row.PSObject.Properties.Match('関連付け用ファイル名')) {
                $errorMessages += "Column '関連付け用ファイル名' not found in CSV file: $($csvFile.FullName)"
                break
            }

            $relatedFileName = $row.'関連付け用ファイル名'.Trim()
            $pdfFilePath = Join-Path -Path $pdfFolderForCsv -ChildPath "$relatedFileName.pdf"
            $txtFilePath = Join-Path -Path $pdfFolderForCsv -ChildPath "$relatedFileName.txt"

            # CSVファイルと関連付けファイル名を出力
            Write-Host "Processing CSV File: $($csvFile.FullName)"
            Write-Host "Related File Name: $relatedFileName"

            # PDFまたはTXTファイルが存在するか確認
            if (-not (Test-Path -Path $pdfFilePath) -and -not (Test-Path -Path $txtFilePath)) {
                $errorMessages += "File not found: $relatedFileName (expected: $pdfFilePath or $txtFilePath)"
                
                if ($errorHandlingFlag -eq 1) {
                    # エラーログを記録し、状態を保存して中断
                    $errorLogFilePath = Join-Path -Path $errorLogFolderPath -ChildPath "$csvFileNameWithoutExtension.log"
                    $errorMessages | Out-File -FilePath $errorLogFilePath -Encoding UTF8
                    Write-Host "Errors found for $csvFileNameWithoutExtension. See log: $errorLogFilePath"
                    $state[$csvFileNameWithoutExtension] = $i
                    $state | ConvertTo-Json | Out-File -FilePath $stateFilePath -Encoding UTF8
                    return
                }
            }
        }

        # エラーメッセージがある場合はログファイルに出力
        if ($errorMessages.Count -gt 0) {
            $errorLogFilePath = Join-Path -Path $errorLogFolderPath -ChildPath "$csvFileNameWithoutExtension.log"
            $errorMessages | Out-File -FilePath $errorLogFilePath -Encoding UTF8
            Write-Host "Errors found for $csvFileNameWithoutExtension. See log: $errorLogFilePath"
            
            if ($errorHandlingFlag -eq 2) {
                $state[$csvFileNameWithoutExtension] = $i
                $state | ConvertTo-Json | Out-File -FilePath $stateFilePath -Encoding UTF8
                return
            }
        } else {
            Write-Host "No errors found for $csvFileNameWithoutExtension."
        }

        # 処理が完了したら状態を削除
        if ($state.ContainsKey($csvFileNameWithoutExtension)) {
            $state.Remove($csvFileNameWithoutExtension)
            $state | ConvertTo-Json | Out-File -FilePath $stateFilePath -Encoding UTF8
        }
    } else {
        Write-Host "Folder not found for CSV: $csvFileNameWithoutExtension"
    }
}