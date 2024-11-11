# 型が利用可能か確認する共通関数
function Test_TypeAvailability {
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

# 必要な型が見つからなかった場合の処理
function Process_MissingTypes {
    Write-Host "必要な型が見つかりませんでした。COMオブジェクトを生成します。"
    # Wordアプリケーションを起動
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # ドキュメントを開く
    $doc = $word.Documents.Open("D:\GitHub\PS_Script\技100-999.docx")
    # WdInformationの直接値を使用
    $wdVerticalPositionRelativeToPage = 1
    $wdHorizontalPositionRelativeToPage = 2
    # 例として、ドキュメントの最初の段落の位置を取得（デバッグ用）
    $position = $doc.Paragraphs[1].Range.Information($wdVerticalPositionRelativeToPage)
    Write-Host "Open Doc Position: $($position)"
    # ドキュメントを保存して閉じる
    $doc.Save()
    $doc.Close()
    $word.Quit()
    # クリーンアップ
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 必要な型が見つかった場合の処理
function Process_AvailableTypes {
    Write-Host "必要な型が見つかりました。処理を続行します。"
    function Find_TablesWithRoles {
        param (
            [__ComObject]$doc,
            [string[]]$roles
        )
        $tablesWithRoles = @()
        if ($doc.Tables.Count -eq 0) {
            Write-Host "ドキュメントにテーブルがありません。"
            return $tablesWithRoles
        }
        foreach ($table in $doc.Tables) {
            if ($table -is [__ComObject]) {
                $foundRoles = @()
                $roleIndex = 0
                foreach ($cell in $table.Range.Cells) {
                    $cellText = $cell.Range.Text.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                        # 不要な文字とスペースを削除
                        $cleanedCellText = $cellText -replace '[^\p{L}\p{N}\p{Zs}]', '' -replace '\s+', ''
                        Write-Host "Cell Text: $($cellText), Cleaned Cell Text: $($cleanedCellText), Role: $($roles[$roleIndex])"
                        if ($cleanedCellText -match [regex]::Escape($roles[$roleIndex])) {
                            $foundRoles += $cell
                            $roleIndex++
                            if ($roleIndex -eq $roles.Length) {
                                break
                            }
                        }
                    }
                }
                if ($foundRoles.Length -eq $roles.Length) {
                    $tablesWithRoles += $table
                }
            }
        }
        return $tablesWithRoles
    }

    function Find_ClosestTableToTopRight {
        param (
            [__ComObject[]]$tables
        )
        $closest_Table = $null
        $closest_To_Top_Right = [double]::MaxValue
        foreach ($table in $tables) {
            $top_Position = $table.Cell(1, 1).Range.Information(1) # Y座標
            $left_Position = $table.Cell(1, 1).Range.Information(2) # X座標
            # 右上角に最も近いテーブルを見つける
            $distance = [math]::Sqrt([math]::Pow($top_Position, 2) + [math]::Pow($left_Position, 2))
            if ($distance -lt $closest_To_Top_Right) {
                $closest_To_Top_Right = $distance
                $closest_Table = $table
            }
        }
        return $closest_Table
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
            $tablesWithRoles = Find_TablesWithRoles -doc $this.Doc -roles $this.Roles
            if ($tablesWithRoles.Length -eq 0) {
                throw "文書内に指定された役割が含まれる表が見つかりませんでした。"
            } elseif ($tablesWithRoles.Length -eq 1) {
                $this.Table = $tablesWithRoles[0]
            } else {
                $this.Table = Find_ClosestTableToTopRight -tables $tablesWithRoles
            }
            $this.Validate_Roles()
        }

        [void] Validate_Roles() {
            $expected_Roles = @("承認", "照査", "作成")
            $table_Roles = @()
            $roleIndex = 0
            foreach ($cell in $this.Table.Range.Cells) {
                $role_Text = $cell.Range.Text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($role_Text)) {
                    # 不要な文字とスペースを削除
                    $cleanedRoleText = $role_Text -replace '[^\p{L}\p{N}\p{Zs}]', '' -replace '\s+', ''
                    Write-Host "Role Text: $($role_Text), Cleaned Role Text: $($cleanedRoleText), Expected Role: $($expected_Roles[$roleIndex])"
                    if ($cleanedRoleText -match [regex]::Escape($expected_Roles[$roleIndex])) {
                        $table_Roles += $role_Text
                        $roleIndex++
                        if ($roleIndex -eq $expected_Roles.Length) {
                            break
                        }
                    }
                }
            }
            Write-Host "Expected Roles: $($expected_Roles)"
            Write-Host "Table Roles: $($table_Roles)"
            if ($expected_Roles.Length -ne $table_Roles.Length) {
                throw "サイン欄フォーマットが違います。役割の文字列が一致しません。期待される役割: $($expected_Roles), 実際の役割: $($table_Roles)"
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

        [array] Get_Cell_Info() {
            $cell_Info = @()

            for ($row = 1; $row -le 2; $row++) {
                for ($col = 1; $col -le 3; $col++) {
                    $cell = $this.Table.Cell($row, $col)
                    $cell_Range = $cell.Range
                    $cell_Text = $cell_Range.Text.Trim()
                    $cell_Position = $cell_Range.Information($this.wdInformation)

                    $cell_Info += [pscustomobject]@{
                        Cell_Number = "R$($row)-C$($col)"
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

                $name_Cell.Range.Text = "$($date_Value)`n$($name_Value)"
                $name_Cell.Range.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
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

                $name_Cell.Range.Text = "$($date)`n$($name)"
                $name_Cell.Range.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
                $name_Cell.Range.Font.Size = 8
                $name_Cell.Range.Paragraphs[2].Range.Font.Size = 10
            }
        }
    }

    # Wordアプリケーションを起動
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # ドキュメントを開く
    $doc = $word.Documents.Open("D:\GitHub\PS_Script\技100-999.docx")
    # 役割配列
    $roles = @("承認", "照査", "作成")
    # Signature_Blockクラスのインスタンスを作成
    try {
        $signature_Block = [Signature_Block]::new($doc, $roles, 1) # 直接値を使用
        Write-Host "Signature_Block インスタンスが正常に作成されました。"
    } catch {
        Write-Error "エラー: $($_)"
        $doc.Close()
        $word.Quit()
        exit 1
    }
    # サイン欄の座標を取得
    $signature_Coordinates = $signature_Block.Get_Signature_Coordinates("left")
    Write-Host "サイン欄原点の座標: $($signature_Coordinates.Origin)"
    Write-Host "サイン欄対角の座標: $($signature_Coordinates.Diagonal)"

    # サイン欄内のセル情報を取得
    $cell_Info = $signature_Block.Get_Cell_Info()
    foreach ($info in $cell_Info) {
        Write-Host "セル番号: $($info.Cell_Number), 位置: $($info.Position), 座標: $($info.Position)"
    }

    # 役割上セルセットの情報を取得
    $role_Above_Cell_Set = $signature_Block.Get_Role_Above_Cell_Set()
    foreach ($cell_Set in $role_Above_Cell_Set) {
        Write-Host "役割: $($cell_Set.Role), 役割セル: $($cell_Set.Role_Cell), 名前セル: $($cell_Set.Name_Cell)"
    }

    # 役割左セルセットの情報を取得
    $role_Left_Cell_Set = $signature_Block.Get_Role_Left_Cell_Set()
    foreach ($cell_Set in $role_Left_Cell_Set) {
        Write-Host "役割: $($cell_Set.Role), 役割セル: $($cell_Set.Role_Cell), 名前セル: $($cell_Set.Name_Cell)"
    }

    # カスタム属性を設定
    $signature_Block.Set_Custom_Attributes()

    # 役割名セルを設定
    $signature_Block.Set_Role_Name_Cells("left")

    # ドキュメントを保存して閉じる
    $doc.Save()
    $doc.Close()
    $word.Quit()
    # クリーンアップ
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# メイン処理部
$wdInformationAvailable = Test_TypeAvailability -typeName "Microsoft.Office.Interop.Word.WdInformation, Microsoft.Office.Interop.Word"
if (-not $wdInformationAvailable) {
    Process_MissingTypes
}
Process_AvailableTypes
