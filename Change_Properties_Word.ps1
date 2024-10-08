function Remove-SealImage {
    Param (
        [Parameter(Mandatory=$true)]
        [System.__ComObject] $Document,
        [Parameter(Mandatory=$true)]
        [string] $ImagePath
    )

    # ドキュメントから印影を削除
    $shapes = $Document.InlineShapes
    for ($i = $shapes.Count; $i -gt 0; $i--) {
        $shape = $shapes.Item($i)
        if ($shape.Type -eq 3) {  # 3 corresponds to wdInlineShapePicture
            $shape.Delete()
        }
    }
}

# . d:\OfficeProperties.ps1

# Microsoft.Office.Interop.Word アセンブリのフルパスを指定してロード
$assemblyPath = "C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\v4.0_15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
Add-Type -Path $assemblyPath

 
write-host "Start Word and load a document..." -Foreground Yellow
$app = New-Object -ComObject Word.Application
$app.visible = $false
$doc = $app.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx", $false, $false, $false)

# 1つ目のテーブルを取得
$table = $doc.Tables.Item(1)

# テーブルのプロパティを取得
$rows = $table.Rows.Count
$columns = $table.Columns.Count

# 各セルの情報を取得

foreach ($row in 1..$rows) {
    foreach ($col in 1..$columns) {
        $cell = $table.Cell($row, $col)
        $cellText = $cell.Range.Text
        Write-host "Row: $row, Column: $col, Text: $cellText"
    }
}
 


# 1つ目のテーブルを取得
# $table = $doc.Tables.Item(1)

# 1つ目のセルを取得
$cell = $table.Cell(2, 6)

# セルの座標とサイズを取得
$left = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdHorizontalPositionRelativeToPage)
$top = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
$width = $cell.Width
$height = $cell.Height

# 画像のサイズを設定
$imageWidth = 50
$imageHeight = 50

# 画像の中央位置を計算
$imageLeft = $left + ($width - $imageWidth) / 2
$imageTop = $top + ($height - $imageHeight) / 2

# 既存の画像を削除（もしあれば）
foreach ($shape in $doc.Shapes) {
    if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdShapeType]::wdInlineShapePicture) {
        $shape.Delete()
    }
}

# 新しい画像を挿入
$shape = $doc.Shapes.AddPicture("C:\Users\y0927\Documents\GitHub\PS_Script\社長印.tif", $false, $true, $imageLeft, $imageTop, $imageWidth, $imageHeight)

# 画像のプロパティを変更
$shape.LockAspectRatio = $false
$shape.Width = 100
$shape.Height = 100




write-host "`nAll BUILT IN Properties:" -Foreground Yellow
Get-OfficeDocBuiltInProperties $doc
 
write-host "`nWrite to BUILT IN author property:" -Foreground Yellow
$result = Set-OfficeDocBuiltInProperty "Author" "Mr. Robot" $doc
write-host "Result: $result"
 
write-host "`nRead BUILT IN author again:" -Foreground Yellow
Get-OfficeDocBuiltInProperty "Author" $doc
 
write-host "`nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc
 
write-host "`nWrite a CUSTOM property:" -Foreground Yellow
# $result = Set-OfficeDocCustomProperty "Batter" "Otani" $doc
$result = Set-OfficeDocCustomProperty "Batter" "小谷" $doc
write-host "Result: $result"

# 印影のパス
$sealImagePath = "C:\Users\y0927\Documents\GitHub\PS_Script\社長印.png"

# カスタムプロパティ "Batter" の値を確認
$batterValue = Get-OfficeDocCustomProperty -PropertyName "Batter" -Document $doc

if ($batterValue -eq "Otani") {
    write-host "`nAdd seal image:" -Foreground Yellow
    Add-SealImage -Document $doc -ImagePath $sealImagePath
} else {
    write-host "`nRemove seal image:" -Foreground Yellow
    Remove-SealImage -Document $doc -ImagePath $sealImagePath
}

write-host "`nRead back the CUSTOM property:" -Foreground Yellow
Get-OfficeDocCustomProperty "Batter" $doc

write-host "`nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc

write-host "`nSave document and close Word..." -Foreground Yellow

# ドキュメントを保存して閉じる
$doc.Save()
$doc.Close()
$word.Quit()

# COMオブジェクトの解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($shape) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

[gc]::collect()
[gc]::WaitForPendingFinalizers()
write-host "`nReady!" -Foreground Green