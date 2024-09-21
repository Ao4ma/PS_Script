param (
    [string]$defaultFilePath = "S:\\技術部storage\\管理課\\PDM復旧\\ファイル.xlsx",
    [string]$homeFolder = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script"
)

function Import-ExcelFile {
    param (
        [string]$filePath = $defaultFilePath
    )

    # エクセルファイルの存在を確認
    if (-Not (Test-Path -Path $filePath)) {
        Write-Error "エクセルファイルが見つかりません: $filePath"
        return
    }

    $outputFolder = Join-Path -Path (Split-Path -Path $filePath -Parent) -ChildPath (Split-Path -Path $filePath -Leaf)
    if (-Not (Test-Path -Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder | Out-Null
    }

    # 出力フォルダ内の既存CSVファイルを削除
    Get-ChildItem -Path $outputFolder -Filter *.csv | Remove-Item -Force

    # エクセルファイルの更新日時を取得
    $excelLastWriteTime = (Get-Item $filePath).LastWriteTime
    $timestampFilePath = Join-Path -Path $outputFolder -ChildPath "timestamp.txt"

    $shouldProcess = $true
    if (Test-Path -Path $timestampFilePath) {
        $lastProcessedTime = [datetime]::Parse((Get-Content -Path $timestampFilePath -Raw))
        if ($excelLastWriteTime -le $lastProcessedTime) {
            $shouldProcess = $false
        }
    }

    if ($shouldProcess) {
        # エクセルファイルをインポート
        try {
            $excel = New-Object -ComObject Excel.Application
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Item(1)
            $usedRange = $sheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $batchSize = 1000
            $batchNumber = 1

            for ($startRow = 1; $startRow -le $rowCount; $startRow += $batchSize) {
                $endRow = [math]::Min($startRow + $batchSize - 1, $rowCount)
                $csvFileName = "{0}_{1:D4}.csv" -f (Split-Path -Path $filePath -Leaf), $batchNumber
                $csvFilePath = Join-Path -Path $outputFolder -ChildPath $csvFileName

                $csvContent = @()
                for ($row = $startRow; $row -le $endRow; $row++) {
                    $rowData = @()
                    foreach ($cell in $sheet.Rows.Item($row).Columns) {
                        $rowData += $cell.Text
                    }
                    $csvContent += ($rowData -join ',')
                }

                $csvContent | Out-File -FilePath $csvFilePath -Encoding UTF8
                Write-Output "CSVファイルが作成されました: $csvFilePath"
                $batchNumber++
            }

            # タイムスタンプを保存
            $excelLastWriteTime.ToString() | Out-File -FilePath $timestampFilePath -Encoding UTF8

            # クリーンアップ
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        } catch {
            Write-Error "エクセルファイルのインポート中にエラーが発生しました: $_"
        }
    } else {
        Write-Output "エクセルファイルに変更がないため、CSVファイルは更新されませんでした。"
    }
}

# デフォルトのエクセルファイルをインポート
Import-ExcelFile