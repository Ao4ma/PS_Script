param (
    [string]$defaultFilePath = "S:\\技術部storage\\管理課\\PDM復旧\\ファイル1.xlsx",
    [string]$homeFolder = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script",
    [int]$batchSize = 0
)

# ExcelProcessorクラスの定義
class ExcelProcessor {
    [string]$FilePath
    [string]$OutputFolder
    [string]$TimestampFilePath
    [int]$MinRowsForCsv = 100

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
}

# デフォルトのエクセルファイルをインポート
$processor = [ExcelProcessor]::new($defaultFilePath)
$processor.ImportExcelFile($batchSize)