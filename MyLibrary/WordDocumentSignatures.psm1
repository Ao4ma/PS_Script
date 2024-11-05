function Fill_Signatures {
    param (
        [WordDocument]$wordDoc,
        [string]$担当者,
        [string]$担当者日付,
        [string]$承認者,
        [string]$承認者日付,
        [string]$照査者,
        [string]$照査者日付
    )
    Write-Host "IN: Fill_Signatures"
    $table = $wordDoc.Document.Tables[1]
    $cell1 = $table.Cell(1, 1)
    $cell2 = $table.Cell(1, 2)
    $cell3 = $table.Cell(1, 3)

    $cell1.Range.Text = "$担当者`n$担当者日付"
    $cell1.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell1.Range.Font.Size = 12
    $cell1.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    $cell2.Range.Text = "$承認者`n$承認者日付"
    $cell2.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell2.Range.Font.Size = 12
    $cell2.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    $cell3.Range.Text = "$照査者`n$照査者日付"
    $cell3.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell3.Range.Font.Size = 12
    $cell3.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    Write-Host "OUT: Fill_Signatures"
}