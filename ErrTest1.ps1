# Microsoft.Office.Interop.Word アセンブリを読み込む
$assemblyPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
try {
    Add-Type -Path $assemblyPath -ErrorAction Stop
    Write-Output "アセンブリが $assemblyPath から正常に読み込まれました"
} catch {
    Write-Error "アセンブリを $assemblyPath から読み込めませんでした。エラー: $_"
    exit 1
}

# 型が利用可能か確認する
if (-not [type]::GetType("Microsoft.Office.Interop.Word.WdInformation, Microsoft.Office.Interop.Word")) {
    Write-Error "型 [Microsoft.Office.Interop.Word.WdInformation] が見つかりません"
    exit 1
}

function Find-SignatureTable {
    param (
        [Microsoft.Office.Interop.Word.Document]$doc
    )

    $highest_Table = $null
    $highest_Position = [double]::MinValue

    foreach ($table in $doc.Tables) {
        if ($table -is [Microsoft.Office.Interop.Word.Table]) {
            $position = $table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            if ($position -gt $highest_Position) {
                $highest_Position = $position
                $highest_Table = $table
            }
        }
    }

    return $highest_Table
}

class Signature_Block {
    $Doc
    [string[]]$Roles
    $Table

    Signature_Block([Microsoft.Office.Interop.Word.Document]$doc, [string[]]$roles) {
        $this.Doc = $doc
        $this.Roles = $roles
        $this.Table = Find-SignatureTable -doc $this.Doc
        if ($null -eq $this.Table -or $this.Table.PSObject.TypeNames -notcontains 'Microsoft.Office.Interop.Word.Table') {
            throw "文書フォーマットが違います。サイン欄の条件に合う表が見つかりませんでした。"
        }
        $this.Validate_Roles()
    }

    [void] Validate_Roles() {
        $expected_Roles = @("承認", "照査", "作成")
        $table_Roles = @()
        for ($col = 1; $col -le 3; $col++) {
            $role_Cell = $this.Table.Cell(1, $col)
            $role_Text = $role_Cell.Range.Text.Trim()
            $table_Roles += $role_Text
        }

        if ($expected_Roles -ne $table_Roles) {
            throw "サイン欄フォーマットが違います。役割の文字列が一致しません。期待される役割: $expected_Roles, 実際の役割: $table_Roles"
        }
    }

    [hashtable] Get_Signature_Coordinates([string]$cell_Set_Type) {
        if ($cell_Set_Type -eq "above") {
            $signature_Origin = $this.Table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            $signature_Diagonal = $this.Table.Cell(2, 3).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
        } elseif ($cell_Set_Type -eq "left") {
            $signature_Origin = $this.Table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            $signature_Diagonal = $this.Table.Cell(1, 3).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
        } else {
            throw "無効なセルセットタイプです。「above」または「left」を使用してください。"
        }
        
        return @{
            Origin = $signature_Origin
            Diagonal = $signature_Diagonal
        }
    }
}

# Wordアプリケーションを起動
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# ドキュメントを開く
$doc = $word.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx")

# 役割配列
$roles = @("承認", "照査", "作成")

# Signature_Blockクラスのインスタンスを作成
try {
    $signature_Block = [Signature_Block]::new($doc, $roles)
    Write-Output "Signature_Block インスタンスが正常に作成されました。"
} catch {
    Write-Error "エラー: $_"
    $doc.Close()
    $word.Quit()
    exit 1
}

# サイン欄の座標を取得
$signature_Coordinates = $signature_Block.Get_Signature_Coordinates("left")
Write-Output "サイン欄原点の座標: $($signature_Coordinates.Origin)"
Write-Output "サイン欄対角の座標: $($signature_Coordinates.Diagonal)"

# ドキュメントを保存して閉じる
$doc.Save()
$doc.Close()
$word.Quit()

# クリーンアップ
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
