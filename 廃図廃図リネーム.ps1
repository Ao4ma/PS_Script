# 修正対象のディレクトリパス
$targetDirectory = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdfデータ_図面-103"

# 指定されたディレクトリ内のテキストファイルを再帰的に取得
$files = Get-ChildItem -Path $targetDirectory -Recurse -Filter "*.txt"

foreach ($file in $files) {
    # ファイル名に「廃図 廃図」が含まれているか確認
    if ($file.Name -like ".*廃図.*廃図.*") {
        # ファイル情報を表示
        Write-Host "Processing file: $($file.FullName)"

        # 新しいファイル名を作成
        $newFileName = $file.Name -replace "廃図 廃図", "廃図"
        $newFilePath = Join-Path -Path $file.DirectoryName -ChildPath $newFileName

        # ファイル名を修正
        Rename-Item -Path $file.FullName -NewName $newFilePath
        Write-Host "Renamed: $($file.FullName) -> $newFilePath"
    }
}