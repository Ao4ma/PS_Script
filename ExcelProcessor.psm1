class ExcelProcessor {
    [string]$FilePath
    [string]$OutputFolder
    [string]$TimestampFilePath
    [int]$MinRowsForCsv = 100
    [System.Collections.ArrayList]$CsvFiles = @()

    # コンストラクタ
    ExcelProcessor([string]$filePath) {
        $this.FilePath = $filePath
        $this.OutputFolder = Join-Path -Path (Split-Path -Path $filePath -Parent) -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($filePath))
        $this.TimestampFilePath = Join-Path -Path $this.OutputFolder -ChildPath "timestamp.txt"
    }

    # 出力フォルダが存在するか確認し、存在しない場合は作成
    [void] EnsureOutputFolderExists() {
        Write-Host "Checking if output folder exists..."
        if (-Not (Test-Path -Path $this.OutputFolder)) {
            Write-Host "Creating output folder..."
            New-Item -ItemType Directory -Path $this.OutputFolder | Out-Null
        }
    }

    # 既存のCSVファイルを削除
    [void] RemoveExistingCsvFiles() {
        Write-Host "Removing existing CSV files..."
        Get-ChildItem -Path $this.OutputFolder -Filter *.csv | Remove-Item -Force
    }

    # ファイルを処理する必要があるか確認
    [bool] ShouldProcessFile() {
        $excelLastWriteTime = (Get-Item $this.FilePath).LastWriteTime
        if (Test-Path -Path $this.TimestampFilePath) {
            $lastProcessedTime = [datetime]::Parse((Get-Content -Path $this.TimestampFilePath -Raw))
            return $excelLastWriteTime -gt $lastProcessedTime
        }
        return $true
    }

    # タイムスタンプを保存
    [void] SaveTimestamp() {
        $excelLastWriteTime = (Get-Item $this.FilePath).LastWriteTime
        $excelLastWriteTime.ToString() | Out-File -FilePath $this.TimestampFilePath -Encoding UTF8
    }

    # エクセルファイルをインポート
    [void] ImportExcelFile([int]$batchSize) {
        try {
            if (-Not (Test-Path -Path $this.FilePath)) {
                throw [System.IO.FileNotFoundException] "エクセルファイルが見つかりません: $this.FilePath"
            }

            $this.EnsureOutputFolderExists()
            $this.RemoveExistingCsvFiles()

            if ($this.ShouldProcessFile()) {
                $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($this.FilePath)
                $fileExtension = [System.IO.Path]::GetExtension($this.FilePath)
                $tempFilePath = Join-Path -Path (Split-Path -Path $this.FilePath -Parent) -ChildPath ($fileNameWithoutExtension + "_副本" + $fileExtension)
                Copy-Item -Path $this.FilePath -Destination $tempFilePath -Force

                $excel = $null
                $workbook = $null
                $sheet = $null
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $workbook = $excel.Workbooks.Open($tempFilePath)
                    $sheet = $workbook.Sheets.Item(1)
                    $usedRange = $sheet.UsedRange
                    $rowCount = $usedRange.Rows.Count
                    $rowCountWithoutTitle = $rowCount - 1

                    $titleRow = @()
                    $columnCount = $sheet.UsedRange.Columns.Count
                    for ($col = 1; $col -le $columnCount; $col++) {
                        $cell = $sheet.Cells.Item(1, $col)
                        $titleRow += $cell.Text
                    }

                    $this.CsvFiles = @()

                    if ($batchSize -eq 0) {
                        $batchSize = $rowCountWithoutTitle
                    }

                    $emptyRowCount = 0
                    $maxEmptyRows = 10  # 連続する空行の最大数

                    for ($startRow = 2; $startRow -le $rowCount; $startRow += $batchSize) {
                        $endRow = [math]::Min($startRow + $batchSize - 1, $rowCount)
                        $batchNumber = [math]::Ceiling(($startRow - 1) / $batchSize)
                        $csvFileName = "{0}_{1:D4}.csv" -f $fileNameWithoutExtension, [int]$batchNumber
                        $csvFilePath = Join-Path -Path $this.OutputFolder -ChildPath $csvFileName

                        $csvContent = @()
                        $csvContent += ($titleRow -join ',')

                        $hasData = $false

                        for ($row = $startRow; $row -le $endRow; $row++) {
                            $rowData = @()
                            for ($col = 1; $col -le $columnCount; $col++) {
                                $cell = $sheet.Cells.Item($row, $col)
                                $rowData += $cell.Text
                            }

                            # 行データがすべて空であるかをチェック
                            if ($rowData -join '' -eq '') {
                                $emptyRowCount++
                                if ($emptyRowCount -ge $maxEmptyRows) {
                                    Write-Host "連続する空行が $maxEmptyRows 行に達したため、処理を終了します。"
                                    break
                                }
                                continue
                            } else {
                                $emptyRowCount = 0
                                $hasData = $true
                            }

                            $csvContent += ($rowData -join ',')
                        }

                        if ($hasData) {
                            # CSVファイルを出力
                            $csvContent | Out-File -FilePath $csvFilePath -Encoding UTF8
                            Write-Host "CSVファイルが作成されました: $csvFilePath"
                            $this.CsvFiles += [PSCustomObject]@{ FileName = $csvFileName; FilePath = $csvFilePath }
                        } else {
                            Write-Host "CSVファイルにデータがないため、作成されませんでした: $csvFilePath"
                        }

                        if ($emptyRowCount -ge $maxEmptyRows) {
                            break
                        }
                    }

                    # タイムスタンプを保存
                    $this.SaveTimestamp()

                } finally {
                    if ($null -ne $workbook) { $workbook.Close($false) }
                    if ($null -ne $excel) { $excel.Quit() }
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    Remove-Item -Path $tempFilePath -Force
                }
            }
        } catch {
            Write-Error "エラーが発生しました: $_"
        }
    }
}

Export-ModuleMember -Class ExcelProcessor