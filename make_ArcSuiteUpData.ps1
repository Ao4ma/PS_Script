$pcName = (hostname).Trim()

# PC名に応じたホームフォルダとワークフォルダの設定
switch ($pcName) {
    "Delld033" {
        $homeFolder = "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\"
        $workFolder = "S:\\技術部storage\\管理課\\PDM復旧"
    }
    "TUF-FX517ZM" {
        $homeFolder = "C:\\Users\\22200\\OneDrive\\ドキュメント\\GitHub\\PS_Script\\"
        $workFolder = "D:\\技術部storage\\管理課\\PDM復旧"
    }
    default {
        throw "Unknown PC name: $pcName"
    }
}
あ
# Ensure the type is available
using module "$homeFolder\ExcelProcessor.psm1"

# モジュールのインポート
Import-Module -Name "$homeFolder\ExcelProcessor.psm1" -ErrorAction Stop

# 特殊文字をエスケープする関数
function Convert-SpecialCharacters {
    param (
        [string]$userInput
    )
    $escapedInput = $userInput -replace '(\[|\])', '`$1'
    return $escapedInput
}

# Excelファイルを処理するクラスの定義
class ExcelHandler {
    [System.Collections.Generic.List[string]] ProcessExcelFile([string]$filePath, [int]$batchSize) {
        $processor = [ExcelProcessor]::new($filePath)
        if (-Not $processor) {
            throw "Unable to instantiate [ExcelProcessor]. Please check the module and type definition."
        }
        $processor.ImportExcelFile($batchSize)

        # CSVファイルが置かれているフォルダをサーチしてCSVファイルのパスを取得
        $csvFolder = $processor.OutputFolder
        $csvFiles = Get-ChildItem -Path $csvFolder -Filter *.csv | Select-Object -ExpandProperty FullName
        return $csvFiles
    }
}

# ファイル操作を行うクラスの定義
class FileManager {
    [void] EnsureFolderExists([string]$folderPath) {
        if (-Not (Test-Path -Path $folderPath)) {
            New-Item -Path $folderPath -ItemType Directory
        } else {
            Remove-Item -Path "$folderPath\*" -Recurse -Force
        }
    }

    [void] CopyFilesBasedOnCsv([string]$csvFilePath, [string]$workFolder, [string]$dataFolder, [string]$realDataFolder, [ref]$successCount, [ref]$failureCount) {
        $csvData = Import-Csv -Path $csvFilePath
        $retryCsvPath = Join-Path -Path $workFolder -ChildPath "retry.csv"

        $recordNumber = 0
        foreach ($record in $csvData) {
            $recordNumber++
            Write-Host "Processing CSV File: $csvFilePath, Record Number: $recordNumber"

            $pcName = $record.PC名
            $fileName = Convert-SpecialCharacters $record.ファイル名
            $fileExtension = Convert-SpecialCharacters $record.拡張子名
            $dateOrder = $record.日付順位
            $fullPath = Convert-SpecialCharacters $record.フルパス

            # SWPDMがつくフォルダを探す
            $swpdmFolder = Get-ChildItem -Path $realDataFolder -Directory -Filter "SWPDM*" | Where-Object { $_.Name -like "*$pcName*" }

            if ($swpdmFolder) {
                $swpdmFolderPath = $swpdmFolder.FullName
                $newFullPath = $fullPath -replace "^C:\\SWPDM", $swpdmFolderPath

                if (Test-Path -Path $newFullPath) {
                    $this.CopyFile($newFullPath, $dataFolder, $dateOrder, $swpdmFolderPath, $workFolder, $record, $successCount, $failureCount)
                } else {
                    $this.HandleFileNotFound($newFullPath, $workFolder, $record, $retryCsvPath, $swpdmFolderPath, $fileName, $fileExtension, $dataFolder, $dateOrder, $successCount, $failureCount)
                }
            } else {
                $this.LogError($workFolder, "SWPDM folder not found for PC: $pcName, File: $fileName.$fileExtension, Full Path: $fullPath, Record: $($record | ConvertTo-Json -Compress)")
                $record | Export-Csv -Path $retryCsvPath -Append -NoTypeInformation
                $failureCount.Value++
            }
        }
    }

    [void] CopyFile([string]$newFullPath, [string]$dataFolder, [string]$dateOrder, [string]$swpdmFolderPath, [string]$workFolder, $record, [ref]$successCount, [ref]$failureCount) {
        $destinationFolder = Join-Path -Path $dataFolder -ChildPath $dateOrder
        if (-Not (Test-Path -Path $destinationFolder)) {
            New-Item -Path $destinationFolder -ItemType Directory
        }

        $swpdmDestinationFolder = Join-Path -Path $destinationFolder -ChildPath "SWPDM"
        if (-Not (Test-Path -Path $swpdmDestinationFolder)) {
            New-Item -Path $swpdmDestinationFolder -ItemType Directory
        }

        $destinationFolder = $swpdmDestinationFolder
        $relativePath = $newFullPath.Substring($swpdmFolderPath.Length)
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath $relativePath
        $destinationDir = Split-Path -Path $destinationPath -Parent

        if (-Not (Test-Path -Path $destinationDir)) {
            New-Item -Path $destinationDir -ItemType Directory -Force
        }

        Copy-Item -Path $newFullPath -Destination $destinationPath -Force
        $successCount.Value++
    }

    [void] HandleFileNotFound([string]$newFullPath, [string]$workFolder, $record, [string]$retryCsvPath, [string]$swpdmFolderPath, [string]$fileName, [string]$fileExtension, [string]$dataFolder, [string]$dateOrder, [ref]$successCount, [ref]$failureCount) {
        $this.LogError($workFolder, "File not found: $newFullPath, Record: $($record | ConvertTo-Json -Compress)")
        $record | Export-Csv -Path $retryCsvPath -Append -NoTypeInformation
        $failureCount.Value++

        $fileToCopy = Get-ChildItem -Path $swpdmFolderPath -Recurse -Filter "$fileName.$fileExtension" | Where-Object {
            $_.FullName -like "*$fullPath"
        } | Select-Object -First 1

        if ($fileToCopy) {
            $this.CopyFile($fileToCopy.FullName, $dataFolder, $dateOrder, $swpdmFolderPath, $workFolder, $record, $successCount, $failureCount)
        } else {
            $this.LogError($workFolder, "File not found after search: $fileName.$fileExtension in $swpdmFolderPath, Record: $($record | ConvertTo-Json -Compress)")
            $record | Export-Csv -Path $retryCsvPath -Append -NoTypeInformation
            $failureCount.Value++
        }
    }

    [void] LogError([string]$workFolder, [string]$message) {
        $errorFile = Join-Path -Path $workFolder -ChildPath "error.log"
        Add-Content -Path $errorFile -Value $message
    }
}

# メイン処理を行うクラスの定義
class MainProcess {
    [void] Run([string]$excelFilePath, [string]$workFolder, [string]$dataFolder, [string]$realDataFolder, [int]$batchSize) {
        $fileManager = [FileManager]::new()
        $fileManager.EnsureFolderExists($dataFolder)

        $logFiles = @("process.log", "error.log", "retry.csv")
        foreach ($logFile in $logFiles) {
            $logFilePath = Join-Path -Path $workFolder -ChildPath $logFile
            if (Test-Path -Path $logFilePath) {
                Remove-Item -Path $logFilePath -Force
            }
        }

        $successCount = [ref]0
        $failureCount = [ref]0

        $excelHandler = [ExcelHandler]::new()
        $csvFilePaths = $excelHandler.ProcessExcelFile($excelFilePath, $batchSize)
        if ($csvFilePaths.Count -eq 0) {
            throw "CSV file paths are empty. Please check the Excel processing."
        }

        foreach ($csvFilePath in $csvFilePaths) {
            $fileManager.CopyFilesBasedOnCsv($csvFilePath, $workFolder, $dataFolder, $realDataFolder, $successCount, $failureCount)
        }

        $logFile = Join-Path -Path $workFolder -ChildPath "process.log"
        $logContent = @(
            "Success Count: $($successCount.Value)",
            "Failure Count: $($failureCount.Value)"
        )
        Add-Content -Path $logFile -Value $logContent
    }
}

# データ保存フォルダ
$dataFolder = Join-Path -Path $workFolder -ChildPath "SWPDM復旧データ"

# 実データフォルダ
$realDataFolder = Join-Path -Path $workFolder -ChildPath "実データ"

# 処理するExcelファイルのパス
$excelFileName = "ファイル.xlsx"
$excelFilePath = Join-Path -Path $workFolder -ChildPath $excelFileName

# バッチサイズ
$batchSize = 5000

# メイン処理の呼び出し
$mainProcess = [MainProcess]::new()
$mainProcess.Run($excelFilePath, $workFolder, $dataFolder, $realDataFolder, $batchSize)

# エラー処理
trap {
    $errorMessage = $_.Exception.Message
    $errorFile = Join-Path -Path $workFolder -ChildPath "error.log"
    Add-Content -Path $errorFile -Value $errorMessage
    continue
}
