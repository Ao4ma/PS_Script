param (
    [string]$folderA = "C:\Path\To\FolderA", # フォルダAのパス
    [string]$folderB = "C:\Path\To\FolderB", # フォルダBのパス
    [string]$logFilePath = $(Join-Path -Path (Split-Path -Path $MyInvocation.MyCommand.Path -Parent) -ChildPath "フォルダ差異確認結果.txt") # ログファイルのパス（デフォルトはスクリプトのある場所）
)

function Get-FolderContent {
    param (
        [string]$folderPath
    )

    # フォルダ内のPDF、TXT、およびTIFFファイルを再帰的に取得
    $items = Get-ChildItem -Path $folderPath -Recurse -Force -Include *.pdf, *.txt, *.tiff | ForEach-Object {
        $_.FullName.Substring($folderPath.Length).TrimStart('\')
    }
    return $items
}

function Get-FolderStats {
    param (
        [string]$folderPath
    )

    $files = Get-ChildItem -Path $folderPath -Recurse -File -Force -Include *.pdf, *.txt, *.tiff
    $folders = Get-ChildItem -Path $folderPath -Recurse -Directory -Force

    $fileCount = $files.Count
    $folderCount = $folders.Count
    $totalSize = ($files | Measure-Object -Property Length -Sum).Sum

    return [PSCustomObject]@{
        FileCount = $fileCount
        FolderCount = $folderCount
        TotalSize = $totalSize
    }
}

# ログファイルの初期化
if (Test-Path -Path $logFilePath) {
    Remove-Item -Path $logFilePath -Force
}
New-Item -Path $logFilePath -ItemType File

# フォルダAとフォルダBの内容を取得
Write-Host "Comparing contents of $folderA and $folderB..."
$contentA = Get-FolderContent -folderPath $folderA
$contentB = Get-FolderContent -folderPath $folderB

# フォルダAにあってフォルダBにないアイテムを取得
$onlyInA = $contentA | Where-Object { $_ -notin $contentB }
# フォルダBにあってフォルダAにないアイテムを取得
$onlyInB = $contentB | Where-Object { $_ -notin $contentA }

# フォルダAとフォルダBの統計情報を取得
Write-Host "Gathering statistics for $folderA and $folderB..."
$statsA = Get-FolderStats -folderPath $folderA
$statsB = Get-FolderStats -folderPath $folderB

# 結果を表示
Write-Host "Items only in $folderA:"
$onlyInA | ForEach-Object { Write-Host $_ }

Write-Host "`nItems only in $folderB:"
$onlyInB | ForEach-Object { Write-Host $_ }

# フォルダAの統計情報を表示
Write-Host "`nStatistics for $folderA:"
Write-Host "File Count: $($statsA.FileCount)"
Write-Host "Folder Count: $($statsA.FolderCount)"
Write-Host "Total Size: $([math]::Round($statsA.TotalSize / 1MB, 2)) MB"

# フォルダBの統計情報を表示
Write-Host "`nStatistics for $folderB:"
Write-Host "File Count: $($statsB.FileCount)"
Write-Host "Folder Count: $($statsB.FolderCount)"
Write-Host "Total Size: $([math]::Round($statsB.TotalSize / 1MB, 2)) MB"

# ログファイルに詳細な結果を出力
Add-Content -Path $logFilePath -Value "Items only in $folderA:"
$onlyInA | ForEach-Object { Add-Content -Path $logFilePath -Value $_ }

Add-Content -Path $logFilePath -Value "`nItems only in $folderB:"
$onlyInB | ForEach-Object { Add-Content -Path $logFilePath -Value $_ }

Add-Content -Path $logFilePath -Value "`nStatistics for $folderA:"
Add-Content -Path $logFilePath -Value "File Count: $($statsA.FileCount)"
Add-Content -Path $logFilePath -Value "Folder Count: $($statsA.FolderCount)"
Add-Content -Path $logFilePath -Value "Total Size: $([math]::Round($statsA.TotalSize / 1MB, 2)) MB"

Add-Content -Path $logFilePath -Value "`nStatistics for $folderB:"
Add-Content -Path $logFilePath -Value "File Count: $($statsB.FileCount)"
Add-Content -Path $logFilePath -Value "Folder Count: $($statsB.FolderCount)"
Add-Content -Path $logFilePath -Value "Total Size: $([math]::Round($statsB.TotalSize / 1MB, 2)) MB"

# エクスプローラーのプロパティと突き合わせる
if ($statsA.FileCount -eq $statsB.FileCount -and $statsA.FolderCount -eq $statsB.FolderCount -and $statsA.TotalSize -eq $statsB.TotalSize) {
    Write-Host "`nThe folder statistics match between $folderA and $folderB."
    Add-Content -Path $logFilePath -Value "`nThe folder statistics match between $folderA and $folderB."
} else {
    Write-Host "`nThe folder statistics do not match between $folderA and $folderB."
    Add-Content -Path $logFilePath -Value "`nThe folder statistics do not match between $folderA and $folderB."
    Add-Content -Path $logFilePath -Value "`nDifferences in folder statistics:"

    if ($statsA.FileCount -ne $statsB.FileCount) {
        $message = "File Count differs: $folderA has $($statsA.FileCount) files, $folderB has $($statsB.FileCount) files."
        Write-Host $message
        Add-Content -Path $logFilePath -Value $message
    }

    if ($statsA.FolderCount -ne $statsB.FolderCount) {
        $message = "Folder Count differs: $folderA has $($statsA.FolderCount) folders, $folderB has $($statsB.FolderCount) folders."
        Write-Host $message
        Add-Content -Path $logFilePath -Value $message
    }

    if ($statsA.TotalSize -ne $statsB.TotalSize) {
        $message = "Total Size differs: $folderA has $([math]::Round($statsA.TotalSize / 1MB, 2)) MB, $folderB has $([math]::Round($statsB.TotalSize / 1MB, 2)) MB."
        Write-Host $message
        Add-Content -Path $logFilePath -Value $message
    }
}