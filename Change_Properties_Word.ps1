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
 
write-host "Start Word and load a document..." -Foreground Yellow
$app = New-Object -ComObject Word.Application
$app.visible = $false
$doc = $app.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx", $false, $false, $false)
 
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
$doc.Save()
$doc.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
write-host "`nReady!" -Foreground Green