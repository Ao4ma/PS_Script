using module .\WordDocument.psm1

class Signature_Block {
    [WordDocument]$WordDoc
    [string[]]$Roles
    [__ComObject]$Table
    $wdInformation 

    Signature_Block([WordDocument]$wordDoc, [string[]]$roles, [int]$wdInformation) {
        $this.WordDoc = $wordDoc
        $this.Roles = $roles
        $this.wdInformation = $wdInformation

        $tablesWithRoles = Find_TablesWithRoles -doc $this.WordDoc.Document -roles $this.Roles
        Write-Host "Tables With Roles: $($tablesWithRoles)"
        Write-Host "Tables With Roles Count: $($tablesWithRoles.Count)"
        if ($tablesWithRoles.Length -eq 0) {
            throw "文書内に指定された役割が含まれる表が見つかりませんでした。"
        } elseif ($tablesWithRoles.Length -eq 2) {
            # $this.Table = $tablesWithRoles[0]
            $this.Table = $tablesWithRoles[1]  # 最初の実際のテーブルを取得
            Write-Host "First Table: $($this.Table)"
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

    [hashtable] Get_Signature_Coordinates() {
        $signature_Info = @{}
        # 役割名のセル座標を取得
        $role_Coordinates = @()
        foreach ($role in $this.Roles) {
            foreach ($cell in $this.Table.Range.Cells) {
                $cellText = $cell.Range.Text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                    # 不要な文字とスペースを削除
                    $cleanedCellText = $cellText -replace '[^\p{L}\p{N}\p{Zs}]', '' -replace '\s+', ''
                    if ($cleanedCellText -eq $role) {
                        $role_Coordinates += [pscustomobject]@{
                            Role = $role
                            Row = $cell.RowIndex
                            Column = $cell.ColumnIndex
                        }
                        break
                    }
                }
            }
        }

        if ($role_Coordinates.Count -ne $this.Roles.Length) {
            throw "役割名のセルが見つかりませんでした。"
        }

        $sign_Cells = @()
        $type = "役割名上タイプ"
        for ($i = 0; $i -lt $role_Coordinates.Count - 1; $i++) {
            if ($role_Coordinates[$i].Column + 1 -ne $role_Coordinates[$i + 1].Column) {
                $type = "役割名左タイプ"
                break
            }
        }

        foreach ($role in $role_Coordinates) {
            if ($type -eq "役割名上タイプ") {
                $sign_Cells += [pscustomobject]@{
                    Role = $role.Role
                    Row = $role.Row + 1
                    Column = $role.Column
                }
            } else {
                $sign_Cells += [pscustomobject]@{
                    Role = $role.Role
                    Row = $role.Row
                    Column = $role.Column + 1
                }
            }
        }

        $signature_Info.Type = $type
        $signature_Info.Sign_Cells = $sign_Cells

        return $signature_Info
    }

    [void] Set_Custom_Attributes_at_signature_Block() {
        $custom_Properties = $this.WordDoc.Document.CustomDocumentProperties
        
        $role_To_Property_Map = @{
            "承認" = @{ Date = "承認日"; Name = "承認者" }
            "照査" = @{ Date = "照査日"; Name = "照査者" }
            "作成" = @{ Date = "作成日"; Name = "作成者" }
        }
        
        $sign_Cells = $this.Get_Signature_Coordinates().Sign_Cells

        foreach ($cell in $sign_Cells) {
            $date_Value = $null
            $name_Value = $null
            $role = $cell.Role
            $row = $cell.Row
            $column = $cell.Column

            $date_Property = $role_To_Property_Map[$role].Date
            $name_Property = $role_To_Property_Map[$role].Name

            if ($null -ne $custom_Properties) {
                $date_Value = $this.WordDoc.Read_Property2($date_Property)
                if ($null -eq $date_Value) {
                    Write-Host "errが発生しました: "
                    $date_Value = "日付なし"
                }
                $name_Value = $this.WordDoc.Read_Property2($name_Property)
                if ($null -eq $name_Value) {
                    Write-Host "errが発生しました: "
                    $name_Value = "日付なし"
                }
            } else {
                $date_Value = "日付なし"
                $name_Value = "名前なし"
            }

            <#
            if ($null -ne $custom_Properties) {
                try {
                    $date_Value = $this.WordDoc.Read_Property2($date_Property)
                } catch {
                    Write-Host "例外が発生しました: $_"
                    $date_Value = "日付なし"
                }

                try {
                    $name_Value = $this.WordDoc.Read_Property2($name_Property)
                } catch {
                    Write-Host "例外が発生しました: $_"
                    $name_Value = "名前なし"
                }
            } else {
                $date_Value = "日付なし"
                $name_Value = "名前なし"
            }

            #>
            Write-Host "書き込み開始"
            $cell_Text = "$($date_Value)`n$($name_Value)"
            $this.Table.Cell($row, $column).Range.Text = $cell_Text
            $this.Table.Cell($row, $column).Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
            $this.Table.Cell($row, $column).Range.Paragraphs[1].Range.Font.Size = 8
            $this.Table.Cell($row, $column).Range.Paragraphs[2].Range.Font.Size = 10
            Write-Host "書き込み完了"
        }
        
        $this.wordDoc.SaveForBugMeasures()
    }




}

function Find_TablesWithRoles {
    param (
        [__ComObject]$doc,
        [string[]]$roles
    )

    $tablesWithRoles = @()

    if ($null -eq $doc.Tables -or $doc.Tables.Count -eq 0) {
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
                $tablesWithRoles += ,$table
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
        $top_Position = $table.Cell(1, 1).Range.Information(1)
        $left_Position = $table.Cell(1, 1).Range.Information(2)

        $distance = [math]::Sqrt([math]::Pow($top_Position, 2) + [math]::Pow($left_Position, 2))
        if ($distance -lt $closest_To_Top_Right) {
            $closest_To_Top_Right = $distance
            $closest_Table = $table
        }
    }

    return $closest_Table
}