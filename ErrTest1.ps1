# 型が利用可能か確認する共通関数
function Test-TypeAvailability {
    param (
        [string]$typeName
    )

    try {
        $type = [type]::GetType($typeName)
        if ($null -eq $type) {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

# Microsoft.Office.Interop.Word アセンブリを読み込む
$assemblyPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
try {
    Add-Type -Path $assemblyPath -ErrorAction Stop
    Write-Output "アセンブリが $assemblyPath から正常に読み込まれました"
} catch {
    Write-Error "アセンブリを $assemblyPath から読み込めませんでした。エラー: $_"
    exit 1
}

# WdInformation と Document の型が利用可能か確認
$wdInformationAvailable = Test-TypeAvailability -typeName "Microsoft.Office.Interop.Word.WdInformation, Microsoft.Office.Interop.Word"
$documentAvailable = Test-TypeAvailability -typeName "Microsoft.Office.Interop.Word.Document, Microsoft.Office.Interop.Word"

if (-not $wdInformationAvailable -or -not $documentAvailable) {
    Write-Output "必要な型が見つかりませんでした。COMオブジェクトを生成します。"

    # Wordアプリケーションを起動
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    # ドキュメントを開く
    $doc = $word.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx")

    # WdInformationの直接値を使用
    $wdVerticalPositionRelativeToPage = 1

    # 例として、ドキュメントの最初の段落の位置を取得（デバッグ用）
    $position = $doc.Paragraphs[1].Range.Information($wdVerticalPositionRelativeToPage)
    Write-Output "Position: $position"

    # ドキュメントを保存して閉じる
    $doc.Save()
    $doc.Close()
    $word.Quit()

    # クリーンアップ
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} else {
    Write-Output "必要な型が見つかりました。"
}

function Find-SignatureTable {
    param (
        [__ComObject]$doc
    )

    $highest_Table = $null
    $highest_Position = [double]::MinValue

    foreach ($table in $doc.Tables) {
        if ($table -is [__ComObject]) {
            $position = $table.Cell(1, 1).Range.Information(1)  # 直接値を使用
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
    $wdInformation 

    Signature_Block([__ComObject]$doc, [string[]]$roles, [int]$wdInformation) {
        $this.Doc = $doc
        $this.Roles = $roles
        $this.wdInformation = $wdInformation
        $this.Table = Find-SignatureTable -doc $this.Doc
        if ($null -eq $this.Table -or $this.Table.PSObject.TypeNames -notcontains '__ComObject') {
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
            $signature_Origin = $this.Table.Cell(1, 1).Range.Information($this.wdInformation)
            $signature_Diagonal = $this.Table.Cell(2, 3).Range.Information($this.wdInformation)
        } elseif ($cell_Set_Type -eq "left") {
            $signature_Origin = $this.Table.Cell(1, 1).Range.Information($this.wdInformation)
            $signature_Diagonal = $this.Table.Cell(1, 3).Range.Information($this.wdInformation)
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
    $signature_Block = [Signature_Block]::new($doc, $roles, 1)  # 直接値を使用
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
