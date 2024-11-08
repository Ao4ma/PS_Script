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
#    $doc = $word.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx")
    $doc = $word.Documents.Open("D:\GitHub\PS_Script\技100-999.docx")

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
   # [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
   # [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
 #   [System.GC]::Collect()
 #   [System.GC]::WaitForPendingFinalizers()
} else {
    Write-Output "必要な型が見つかりました。処理を続行します。"

    function Find-SignatureTable {
        param (
            [Microsoft.Office.Interop.Word.Document]$doc
        )

        $highest_Table = $null
        $highest_Position = [double]::MinValue

        for ($i = 1; $i -le $doc.Tables.Count; $i++) {
            $table = $doc.Tables.Item($i)
            if ($table -is [Microsoft.Office.Interop.Word.Table]) {
                $topLeft = $table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
                $bottomRight = $table.Cell($table.Rows.Count, $table.Columns.Count).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
                Write-Output "Table Index: $i, Top-Left Position: $topLeft, Bottom-Right Position: $bottomRight, Rows: $($table.Rows.Count), Columns: $($table.Columns.Count)"
            } else {
                Write-Output "Table Index: $i is not a valid Table object."
            }
        }

        foreach ($table in $doc.Tables) {
            if ($table -is [Microsoft.Office.Interop.Word.Table]) {
                $position = $table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
                Write-Output "Table Position: $position"
                if ($position -gt $highest_Position) {
                    $highest_Position = $position
                    $highest_Table = $table
                    Write-Output "New Highest Table Found: Position = $highest_Position"
                }
            } else {
                Write-Output "Invalid Table object encountered."
            }
        }

        Write-Output "Highest Table Position: $highest_Position"
        if ($highest_Table -is [Microsoft.Office.Interop.Word.Table]) {
            Write-Output "Highest Table Object Type: $($highest_Table.GetType().FullName)"
        } else {
            Write-Output "Highest Table is not a valid Table object."
        }
        return $highest_Table
    }

    class Signature_Block {
        $Doc
        [string[]]$Roles
        $Table

        Signature_Block($doc, [string[]]$roles) {
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
            Write-Output "Expected Roles: $expected_Roles"
            Write-Output "Table Roles: $table_Roles"

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
                throw "Invalid cell set type. Use 'above' or 'left'."
            }
            
            return @{
                Origin = $signature_Origin
                Diagonal = $signature_Diagonal
            }
        }
        [array] Get_Cell_Info() {
            $cell_Info = @()
            
            for ($row = 1; $row -le 2; $row++) {
                for ($col = 1; $col -le 3; $col++) {
                    $cell = $this.Table.Cell($row, $col)
                    $cell_Range = $cell.Range
                    $cell_Text = $cell_Range.Text.Trim()
                    $cell_Position = $cell_Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
                    
                    $cell_Info += [pscustomobject]@{
                        Cell_Number = "R$row-C$col"
                        Text = $cell_Text
                        Position = $cell_Position
                    }
                }
            }
            
            return $cell_Info
        }
    
        [void] Set_Custom_Attributes() {
            $custom_Properties = $this.Doc.CustomDocumentProperties
    
            $role_To_Property_Map = @{
                "承認" = @{ Date = "承認日"; Name = "承認者" }
                "照査" = @{ Date = "照査日"; Name = "照査者" }
                "作成" = @{ Date = "作成日"; Name = "作成者" }
            }
    
            foreach ($role in $this.Roles) {
                $role_Index = $this.Roles.IndexOf($role) + 1
                $name_Cell = $this.Table.Cell(2, $role_Index)
    
                $date_Property = $role_To_Property_Map[$role].Date
                $name_Property = $role_To_Property_Map[$role].Name
    
                $date_Value = $custom_Properties.Item($date_Property).Value
                $name_Value = $custom_Properties.Item($name_Property).Value
    
                $name_Cell.Range.Text = "$date_Value`n$name_Value"
                $name_Cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                $name_Cell.Range.Paragraphs[1].Range.Font.Size = 8
                $name_Cell.Range.Paragraphs[2].Range.Font.Size = 10
            }
        }
    
        [array] Get_Role_Above_Cell_Set() {
            $role_Above_Cell_Set = @()
            
            foreach ($role in $this.Roles) {
                $role_Cell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 1))
                $name_Cell = $this.Table.Cell(2, ($this.Roles.IndexOf($role) + 1))
                
                $role_Above_Cell_Set += [pscustomobject]@{
                    Role = $role
                    Role_Cell = $role_Cell
                    Name_Cell = $name_Cell
                }
            }
            
            return $role_Above_Cell_Set
        }
    
        [array] Get_Role_Left_Cell_Set() {
            $role_Left_Cell_Set = @()
            
            foreach ($role in $this.Roles) {
                $role_Cell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 1))
                $name_Cell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 2))
                
                $role_Left_Cell_Set += [pscustomobject]@{
                    Role = $role
                    Role_Cell = $role_Cell
                    Name_Cell = $name_Cell
                }
            }
            
            return $role_Left_Cell_Set
        }
    
        [void] Set_Role_Name_Cells([string]$cell_Set_Type) {
            if ($cell_Set_Type -eq "above") {
                $role_Cell_Set = $this.Get_Role_Above_Cell_Set()
            } elseif ($cell_Set_Type -eq "left") {
                $role_Cell_Set = $this.Get_Role_Left_Cell_Set()
            } else {
                throw "Invalid cell set type. Use 'above' or 'left'."
            }
    
            foreach ($cell_Set in $role_Cell_Set) {
                $role = $cell_Set.Role
                $name_Cell = $cell_Set.Name_Cell
                $date = Get-Date -Format "yyyy/MM/dd"
                $name = "名前"
                
                $name_Cell.Range.Text = "$date`n$name"
                $name_Cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                $name_Cell.Range.Font.Size = 8
                $name_Cell.Range.Paragraphs[2].Range.Font.Size = 10
            }
        }    }

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
        Write-Output "Signature_Block instance created successfully."
    } catch {
        Write-Error "エラー: $_"
        $doc.Close()
        $word.Quit()
        exit 1
    }

    # サイン欄の座標を取得
    $signature_Coordinates = $signature_Block.Get_Signature_Coordinates("left")
    Write-Output "サイン欄原点の座標: $($signature_Coordinates.Origin)"
    Write-Output "サイン欄対角の座標: $($signature_Coordinates.Diagonal)"# サイン欄内のセル情報を取得
    $cell_Info = $signature_Block.Get_Cell_Info()
    $cell_Info | ForEach-Object { Write-Output "セル番号: $($_.Cell_Number), 位置: $($_.Position), 座標: $($_.Position)" }
    
    # 役割上セルセットの情報を取得
    $role_Above_Cell_Set = $signature_Block.Get_Role_Above_Cell_Set()
    $role_Above_Cell_Set | ForEach-Object { Write-Output "役割: $($_.Role), 役割セル: $($_.Role_Cell), 名前セル: $($_.Name_Cell)" }
    
    # 役割左セルセットの情報を取得
    $role_Left_Cell_Set = $signature_Block.Get_Role_Left_Cell_Set()
    $role_Left_Cell_Set | ForEach-Object { Write-Output "役割: $($_.Role), 役割セル: $($_.Role_Cell), 名前セル: $($_.Name_Cell)" }
    
    # カスタム属性を設定
    $signature_Block.Set_Custom_Attributes()
    
    # 役割名セルを設定
    $signature_Block.Set_Role_Name_Cells("left")
    
    # ドキュメントを保存して閉じる
    $doc.Save()
    $doc.Close()
    $word.Quit()

    # クリーンアップ
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}