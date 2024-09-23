# モジュールのインポート
using module "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\ExcelProcessor.psm1"
Import-Module -Name "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\ExcelProcessor.psm1" -ErrorAction Stop

# Verify if the type [ExcelProcessor.ExcelProcessor] is available
if (-Not ([type]::GetType("ExcelProcessor.ExcelProcessor"))) {
    throw "Unable to find type [ExcelProcessor.ExcelProcessor]. Please check the module path and type name."
}

# 特殊文字をエスケープする関数
function Convert-SpecialCharacters {
    param (
        [string]$inputString
    )
    $escapedInput = $inputString -replace '([\\\*\?\|\<\>\:\"]|\[|\])', '\\$1'
    return $escapedInput
}

# Excelファイルを処理するクラスの定義
class ExcelHandler {
    [System.Collections.Generic.List[string]] ProcessExcelFile([string]$filePath, [int]$batchSize) {
        $processor = [ExcelProcessor.ExcelProcessor]::new($filePath)
        if (-Not $processor) {
            throw "Unable to instantiate [ExcelProcessor.ExcelProcessor]. Please check the module and type definition."
        }
        if (-Not $processor) {
            throw "Unable to instantiate [ExcelProcessor.ExcelProcessor]. Please check the module and type definition."
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

    [void] CopyFilesBasedOnCsv([string]$csvFilePath, [string]$workFolder, [string]$dataFolder, [string]$realDataFolder) {
        $csvData = Import-Csv -Path $csvFilePath

        foreach ($record in $csvData) {
            $pcName = $record.PC名
            $fileName = Escape-SpecialCharacters $record.ファイル名
            $fileExtension = Escape-SpecialCharacters $record.拡張子名
            $index = $record.インデックス
            $fullPath = Escape-SpecialCharacters $record.フルパス

            # SWPDMがつくフォルダを探す
            $swpdmFolder = Get-ChildItem -Path $realDataFolder -Directory -Filter "SWPDM*" | Where-Object { $_.Name -like "*$pcName*" }

            if ($swpdmFolder) {
                $swpdmFolderPath = $swpdmFolder.FullName
                $newFullPath = $fullPath -replace "^C:\\SWPDM", $swpdmFolderPath

                if (Test-Path -Path $newFullPath) {
                    # コピー先フォルダを作成
                    $destinationFolder = Join-Path -Path $dataFolder -ChildPath $index
                    if (-Not (Test-Path -Path $destinationFolder)) {
                        New-Item -Path $destinationFolder -ItemType Directory
                    }

                    # SWPDMフォルダを作成
                    $swpdmDestinationFolder = Join-Path -Path $destinationFolder -ChildPath "SWPDM"
                    if (-Not (Test-Path -Path $swpdmDestinationFolder)) {
                        New-Item -Path $swpdmDestinationFolder -ItemType Directory
                    }

                    # destinationFolderをSWPDMフォルダに設定
                    $destinationFolder = $swpdmDestinationFolder

                    # フォルダ構成を維持してファイルをコピー
                    $relativePath = $newFullPath.Substring($swpdmFolderPath.Length)
                    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $relativePath
                    $destinationDir = Split-Path -Path $destinationPath -Parent

                    if (-Not (Test-Path -Path $destinationDir)) {
                        New-Item -Path $destinationDir -ItemType Directory -Force
                    }

                    Copy-Item -Path $newFullPath -Destination $destinationPath -Force
                } else {
                    # エラーログに記録
                    $errorFile = Join-Path -Path $workFolder -ChildPath "error.log"
                    Add-Content -Path $errorFile -Value "File not found: $newFullPath"

                    # ファイルを再帰的にサーチ
                    $fileToCopy = Get-ChildItem -Path $swpdmFolderPath -Recurse -Filter "$fileName.$fileExtension" | Where-Object {
                        $_.FullName -like "*$fullPath"
                    } | Select-Object -First 1

                    if ($fileToCopy) {
                        # コピー先フォルダを作成
                        $destinationFolder = Join-Path -Path $dataFolder -ChildPath $index
                        if (-Not (Test-Path -Path $destinationFolder)) {
                            New-Item -Path $destinationFolder -ItemType Directory
                        }

                        # SWPDMフォルダを作成
                        $swpdmDestinationFolder = Join-Path -Path $destinationFolder -ChildPath "SWPDM"
                        if (-Not (Test-Path -Path $swpdmDestinationFolder)) {
                            New-Item -Path $swpdmDestinationFolder -ItemType Directory
                        }

                        # destinationFolderをSWPDMフォルダに設定
                        $destinationFolder = $swpdmDestinationFolder

                        # フォルダ構成を維持してファイルをコピー
                        $relativePath = $fileToCopy.FullName.Substring($swpdmFolderPath.Length)
                        $destinationPath = Join-Path -Path $destinationFolder -ChildPath $relativePath
                        $destinationDir = Split-Path -Path $destinationPath -Parent

                        if (-Not (Test-Path -Path $destinationDir)) {
                            New-Item -Path $destinationDir -ItemType Directory -Force
                        }

                        Copy-Item -Path $fileToCopy.FullName -Destination $destinationPath -Force
                    } else {
                        # エラーログに記録
                        Add-Content -Path $errorFile -Value "File not found after search: $fileName.$fileExtension in $swpdmFolderPath"
                    }
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
        [string]$realDataFolder,
        [int]$batchSize
    )

    # フォルダが存在しない場合は作成
    $fileManager = [FileManager]::new()
    $fileManager.EnsureFolderExists($dataFolder)

    # Excelファイルを処理
    $excelHandler = [ExcelHandler]::new()
    $csvFilePaths = $excelHandler.ProcessExcelFile($excelFilePath, $batchSize)
    if ($csvFilePaths.Count -eq 0) {
        throw "CSV file paths are empty. Please check the Excel processing."
    }

    # CSVファイルに基づいてファイルをコピー
    foreach ($csvFilePath in $csvFilePaths) {
        $fileManager.CopyFilesBasedOnCsv($csvFilePath, $workFolder, $dataFolder, $realDataFolder)
    }
}

# ワーク場所のトップフォルダ
$workFolder = "S:\\技術部storage\\管理課\\PDM復旧"

# データ保存フォルダ
$dataFolder = Join-Path -Path $workFolder -ChildPath "SWPDM復旧データ"

# 実データフォルダ
$realDataFolder = Join-Path -Path $workFolder -ChildPath "実データ"

# 処理するExcelファイルのパス
$excelFileName = "ファイル1.xlsx"
$excelFilePath = Join-Path -Path $workFolder -ChildPath $excelFileName

# バッチサイズ
$batchSize = 1000

# メイン処理の呼び出し
Main -excelFilePath $excelFilePath -workFolder $workFolder -dataFolder $dataFolder -realDataFolder $realDataFolder -batchSize $batchSize

# エラー処理
trap {
    $errorMessage = $_.Exception.Message
    $errorFile = Join-Path -Path $workFolder -ChildPath "error.log"
    Add-Content -Path $errorFile -Value $errorMessage
    continue
}