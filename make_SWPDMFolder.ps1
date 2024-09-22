# エクセルファイルをインポート
function ImportExcelFile {
    param (
        [int]$batchSize
    )
    if (-Not (Test-Path -Path $this.FilePath)) {
        Write-Error "エクセルファイルが見つかりません: $this.FilePath"
        return
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

            $csvFiles = @()

            if ($batchSize -eq 0) {
                $batchSize = $rowCountWithoutTitle
            }

            for ($startRow = 2; $startRow -le $rowCount; $startRow += $batchSize) {
                $endRow = [math]::Min($startRow + $batchSize - 1, $rowCount)
                $batchNumber = [math]::Ceiling(($startRow - 1) / $batchSize)
                $csvFileName = "{0}_{1:D4}.csv" -f $fileNameWithoutExtension, [int]$batchNumber
                $csvFilePath = Join-Path -Path $this.OutputFolder -ChildPath $csvFileName

                $csvContent = @()
                $csvContent += ($titleRow -join ',')

                for ($row = $startRow; $row -le $endRow; $row++) {
                    $rowData = @()
                    for ($col = 1; $col -le $columnCount; $col++) {
                        $cell = $sheet.Cells.Item($row, $col)
                        $rowData += $cell.Text
                    }

                    # 一列目が空文字列の場合、取り込みを中断
                    if ($rowData[0] -eq "") {
                        Write-Host "一列目が空文字列のため、取り込みを中断します: Row $row"
                        return
                    }

                    $csvContent += ($rowData -join ',')
                }

                if ($csvContent.Count -ge $this.MinRowsForCsv + 1) {
                    $csvContent | Out-File -FilePath $csvFilePath -Encoding UTF8
                    Write-Host "CSVファイルが作成されました: $csvFilePath"
                    $csvFiles += [PSCustomObject]@{ FileName = $csvFileName; FilePath = $csvFilePath }
                } else {
                    Write-Host "CSVファイルの行数が100未満のため、作成されませんでした: $csvFilePath"
                }
            }

            # タイムスタンプを保存
            $this.SaveTimestamp()

            # CSVファイル情報を保存
            $this.CsvFiles = $csvFiles
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Remove-Item -Path $tempFilePath -Force
        }
    }
}
