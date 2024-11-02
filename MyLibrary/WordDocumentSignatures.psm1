# MyLibrary/WordDocumentSignatures.psm1
function FillSignatures {
    param ($this)
    Write-Host "IN: FillSignatures"

    # カスタムプロパティから担当者、承認者、照査者の名前と日付を取得
    $担当者 = $this.Read_Property("担当者")
    $担当者日付 = $this.Read_Property("担当者日付")
    $承認者 = $this.Read_Property("承認者")
    $承認者日付 = $this.Read_Property("承認者日付")
    $照査者 = $this.Read_Property("照査者")
    $照査者日付 = $this.Read_Property("照査者日付")

    # フォントサイズを計算する関数
    function CalculateFontSize {
        param ($text, $cellWidth)
        $averageCharWidth = 0.6 # 平均的な文字の幅の割合（文字数に基づく調整）
        $textLength = $text.Length
        $fontSize = [math]::Floor($cellWidth / ($averageCharWidth * $textLength))
        return $fontSize - 1 # 少し余裕を持たせるために1ポイント減らす
    }

    # 表のサイン欄に名前と日付を配置
    $table = $this.Document.Tables.Item(1) # 1番目の表を取得
    $cell1 = $table.Cell(2, 1) # 2行1列目
    $cell2 = $table.Cell(2, 2) # 2行2列目
    $cell3 = $table.Cell(2, 3) # 2行3列目

    # サイン欄の横幅を取得
    $cellWidth1 = $cell1.Width
    $cellWidth2 = $cell2.Width
    $cellWidth3 = $cell3.Width

    # フォントサイズを計算
    $fontSize1 = CalculateFontSize $担当者 $cellWidth1
    $fontSize2 = CalculateFontSize $承認者 $cellWidth2
    $fontSize3 = CalculateFontSize $照査者 $cellWidth3

    # 担当者のサイン欄
    $cell1.Range.Text = "$担当者`n$担当者日付"
    $cell1.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell1.Range.Font.Size = $fontSize1
    $cell1.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    # 承認者のサイン欄
    $cell2.Range.Text = "$承認者`n$承認者日付"
    $cell2.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell2.Range.Font.Size = $fontSize2
    $cell2.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    # 照査者のサイン欄
    $cell3.Range.Text = "$照査者`n$照査者日付"
    $cell3.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $cell3.Range.Font.Size = $fontSize3
    $cell3.Range.Paragraphs[2].Range.Font.Size = 8 # 日付のフォントサイズを8に設定

    Write-Host "OUT: FillSignatures"
}