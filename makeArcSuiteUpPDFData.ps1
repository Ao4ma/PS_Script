# 
# 概要:
# このスクリプトは、CSVファイルからデータを読み込み、ファイルデータを表すクラスのインスタンスを
# ハッシュテーブルに格納します。
#
# クラス:
# - FileData: ファイル名とファイルパスを保持するクラス。
#
# 関数:
# - Import-FileDataFromCsv: 指定されたCSVファイルからデータを読み込み、FileDataクラスのインスタンスを
#   ハッシュテーブルに格納します。
#
# パラメータ:
# - csvPath: 読み込むCSVファイルのパスを指定します。
#
# 使用例:
# $csvPath = "C:\path\to\your\file.csv"
# $fileDataHashTable = Import-FileDataFromCsv -csvPath $csvPath
#
# ハッシュテーブルの内容を確認するための出力例:
# $fileDataHashTable.GetEnumerator() | ForEach-Object {
#     Write-Output "FileName: $($_.Key), FilePath: $($_.Value.FilePath)"
# }
# Define a class to represent the file data
class FileData {
    [string]$FileName
    [string]$FilePath

    FileData([string]$fileName, [string]$filePath) {
        $this.FileName = $fileName
        $this.FilePath = $filePath
    }
}

# Function to read CSV and store data in a hash table
function Import-FileDataFromCsv {
    param (
        [string]$csvPath
    )

    $fileDataHashTable = @{}

    # Import CSV data
    $csvData = Import-Csv -Path $csvPath

    foreach ($row in $csvData) {
        $fileData = [FileData]::new($row.FileName, $row.FilePath)
        $fileDataHashTable[$fileData.FileName] = $fileData
    }

    return $fileDataHashTable
}

# Example usage
$csvPath = "C:\path\to\your\file.csv"
$fileDataHashTable = Import-FileDataFromCsv -csvPath $csvPath

# Output the hash table for verification
$fileDataHashTable.GetEnumerator() | ForEach-Object {
    Write-Output "FileName: $($_.Key), FilePath: $($_.Value.FilePath)"
}