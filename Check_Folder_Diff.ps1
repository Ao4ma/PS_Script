param (
    [string]$folderA = "C:\Path\To\FolderA", # フォルダAのパス
    [string]$folderB = "C:\Path\To\FolderB"  # フォルダBのパス
)

function Get-FolderContent {
    param (
        [string]$folderPath
    )

    # フォルダ内のファイルとサブフォルダを再帰的に取得
    $items = Get-ChildItem -Path $folderPath -Recurse | ForEach-Object {
        $_.FullName.Substring($folderPath.Length)
    }
    return $items
}

# フォルダAとフォルダBの内容を取得
$contentA = Get-FolderContent -folderPath $folderA
$contentB = Get-FolderContent -folderPath $folderB

# フォルダAにあってフォルダBにないアイテムを取得
$onlyInA = $contentA | Where-Object { $_ -notin $contentB }
# フォルダBにあってフォルダAにないアイテムを取得
$onlyInB = $contentB | Where-Object { $_ -notin $contentA }

# 結果を表示
Write-Output "Items only in $folderA:"
$onlyInA | ForEach-Object { Write-Output $_ }

Write-Output "`nItems only in $folderB:"
$onlyInB | ForEach-Object { Write-Output $_ }