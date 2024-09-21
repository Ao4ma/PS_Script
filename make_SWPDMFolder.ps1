param (
    [string]$defaultFilePath = "S:\\技術部storage\\管理課\\PDM復旧\\ファイル1.xlsx",
    [string]$homeFolder = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script"
)

class ExcelProcessor {
    [string]$FilePath
    [string]$OutputFolder
    [string]$TimestampFilePath

    ExcelProcessor([string]$filePath) {
        $this.FilePath = $filePath
        $this.OutputFolder = Join-Path -Path (Split-Path -Path $filePath -Parent) -ChildPath (Split-Path -Path $filePath -Leaf)
        $this.TimestampFilePath = Join-Path -Path $this.OutputFolder -ChildPath "timestamp.txt"
    }

    [void] EnsureOutputFolderExists() {
        if (-Not (Test-Path -Path $this.OutputFolder)) {
            New-Item -ItemType Directory -Path $this.OutputFolder | Out-Null
        }
    }

    [void] RemoveExistingCsvFiles() {
        Get-ChildItem -Path $this.OutputFolder -Filter *.csv | Remove-Item -Force
    }

    [bool] ShouldProcessFile() {
        $excelLastWriteTime = (Get-Item $this.FilePath).LastWriteTime
        if (Test-Path -Path $this.TimestampFilePath) {
            $lastProcessedTime = [datetime]::Parse((Get-Content -Path $this.TimestampFilePath -Raw))
            return $excelLastWriteTime -gt $lastProcessedTime
        }
        return $true
    }

    [void] SaveTimestamp() {
        $excelLastWriteTime = (Get-Item $this.FilePath).LastWriteTime
        $excelLastWriteTime.ToString() | Out-File -FilePath $this.TimestampFilePath -Encoding UTF8
    }

    [void] ImportExcelFile() {
        if (-Not (Test-Path -Path $this.FilePath)) {
            Write-Error "エクセルファイルが見つかりません: $this.FilePath"
            return
        }

        $this.EnsureOutputFolderExists()
        $this.RemoveExistingCsvFiles()

        if ($this.ShouldProcessFile()) {
            $tempFilePath = Join-Path -Path (Split-Path -Path $this.FilePath -Parent) -ChildPath ((Split-Path -Path $this.FilePath -LeafBaseName) + "_副本" + (Split-Path -Path $this.FilePath -Extension))
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
                $batchSize = 1000
                $batchNumber = 1

                for ($startRow = 1; $startRow -le $rowCount; $startRow += $batchSize) {
                    $endRow = [math]::Min($startRow + $batchSize - 1, $rowCount)
                    $csvFileName = "{0}_{1:D4}.csv" -f (Split-Path -Path $this.FilePath -LeafBaseName), $batchNumber
                    $csvFilePath = Join-Path -Path $this.OutputFolder -ChildPath $csvFileName

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

                $this.SaveTimestamp()

            } catch {
                Write-Error "エクセルファイルのインポート中にエラーが発生しました: $_"
            } finally {
                if ($null -ne $workbook) {
                    $workbook.Close($false)
                }
                if ($null -ne $excel) {
                    $excel.Quit()
                    if ($null -ne $sheet) {
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
                    }
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                }
            }
        } else {
            Write-Output "エクセルファイルに変更がないため、CSVファイルは更新されませんでした。"
        }
    }
}

# Create an instance of ExcelProcessor and import the default Excel file
$processor = [ExcelProcessor]::new($defaultFilePath)
$processor.ImportExcelFile()
