# 対象フォルダのパス
$targetFolderPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdf_生成場所"

# 対象のファイル名リスト
$fileNames = @(
    "0007f8b9CK-796",
    "0003fc00880668C2",
    "00081312771742B3",
    "000637da853782C2",
    "000400ca880406D2-1",
    "0003eedc880406C2",
    "0006e1c8880519E2",
    "0003fb59880519C2",
    "00040050880519D2",
    "0003eedd880520C2",
    "000400cb880520C2-1",
    "00087498取説No2762"
)

# フォルダ内のPDFファイルを取得
$pdfFiles = Get-ChildItem -Path $targetFolderPath -Recurse -Filter *.pdf

foreach ($pdfFile in $pdfFiles) {
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($pdfFile.Name)
    $fileExtension = $pdfFile.Extension

    # ファイル名がリストに含まれているかチェック
    if ($fileNames -contains $fileNameWithoutExtension) {
        # ファイル名に半角スペースが含まれているかチェック
        if ($fileNameWithoutExtension -match "\s") {
            $newFileNameWithoutExtension = $fileNameWithoutExtension -replace "\s", ""
            $newFileName = "$newFileNameWithoutExtension$fileExtension"
            $newFilePath = Join-Path -Path $targetFolderPath -ChildPath $newFileName

            # 確認ダイアログを表示
            $result = [System.Windows.Forms.MessageBox]::Show("Rename $($pdfFile.FullName) to $newFilePath?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)

            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # ファイル名を変更
                Rename-Item -Path $pdfFile.FullName -NewName $newFileName

                # 結果を出力
                Write-Host "Renamed: $($pdfFile.FullName) -> $newFilePath"
            } else {
                Write-Host "Skipped: $($pdfFile.FullName)"
            }
        } else {
            Write-Host "No space found in: $($pdfFile.FullName)"
        }
    } else {
        Write-Host "File not in list: $($pdfFile.FullName)"
    }
}
