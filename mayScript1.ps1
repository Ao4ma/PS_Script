# モジュールのインポート
Import-Module -Name "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\ExcelProcessor.psm1"

# Excelファイルを処理するクラスの定義
class ExcelHandler {
    [string] ProcessExcelFile([string]$filePath, [int]$batchSize) {
        $processor = [ExcelProcessor]::new($filePath)
        $processor.ImportExcelFile($batchSize)
        return $processor.OutputCsvFilePath  # 生成されたCSVファイルのパスを返す
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

    [void] CopyFilesBasedOnCsv([string]$csvFilePath, [string]$workFolder, [string]$dataFolder) {
        $csvData = Import-Csv -Path $csvFilePath

        foreach ($record in $csvData) {
            $pcName = $record.PC名
            $fileName = $record.ファイル名
            $fileExtension = $record.拡張子名
            $index = $record.インデックス
            $fullPath = $record.フルパス

            # SWPDMがつくフォルダを探す
            $swpdmFolder = Get-ChildItem -Path $workFolder -Directory -Filter "SWPDM*" | Where-Object { $_.Name -like "*$pcName*" }

            if ($swpdmFolder) {
                # ファイルを検索する
                $fileToCopy = Get-ChildItem -Path $swpdmFolder.FullName -Recurse -Filter "$fileName.$fileExtension" | Where-Object {
                    $_.FullName -like "*$fullPath"
                } | Select-Object -First 1

                if ($fileToCopy) {
                    # コピー先フォルダを作成
                    $destinationFolder = Join-Path -Path $dataFolder -ChildPath $index
                    if (-Not (Test-Path -Path $destinationFolder)) {
                        New-Item -Path $destinationFolder -ItemType Directory
                    }

                    # フォルダ構成を維持してファイルをコピー
                    $relativePath = $fileToCopy.FullName.Substring($swpdmFolder.FullName.Length)
                    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $relativePath
                    $destinationDir = Split-Path -Path $destinationPath -Parent

                    if (-Not (Test-Path -Path $destinationDir)) {
                        New-Item -Path $destinationDir -ItemType Directory -Force
                    }

                    Copy-Item -Path $fileToCopy.FullName -Destination $destinationPath -Force
                }
            }
        }
    }
}

# メイン処理
function Main {
    param (
        [string]$excelFilePath,
        [string]$workFolder,
        [string]$dataFolder,
        [int]$batchSize
    )

    # フォルダが存在しない場合は作成
    $fileManager = [FileManager]::new()
    $fileManager.EnsureFolderExists($dataFolder)

    # Excelファイルを処理
    $excelHandler = [ExcelHandler]::new()
    $csvFilePath = $excelHandler.ProcessExcelFile($excelFilePath, $batchSize)

    # CSVファイルに基づいてファイルをコピー
    $fileManager.CopyFilesBasedOnCsv($csvFilePath, $workFolder, $dataFolder)
}

# ワーク場所のトップフォルダ
$workFolder = "S:\\技術部storage\\管理課\\PDM復旧"

# データ保存フォルダ
$dataFolder = Join-Path -Path $workFolder -ChildPath "SWPDM復旧データ"

# 処理するExcelファイルのパス
$excelFileName = "ファイル1.xlsx"
$excelFilePath = Join-Path -Path $workFolder -ChildPath $excelFileName

# バッチサイズ
$batchSize = 100

# メイン処理の呼び出し
Main -excelFilePath $excelFilePath -workFolder $workFolder -dataFolder $dataFolder -batchSize $batchSize

# エラー処理
trap {
    $errorMessage = $_.Exception.Message
    $errorFile = Join-Path -Path $workFolder -ChildPath "error.log"
    Add-Content -Path $errorFile -Value $errorMessage
    continue
}
