using module ".\MyLibrary\WordDocumentProperties.psm1"
using module ".\MyLibrary\WordDocument.psm1"

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

# デバッグ用設定
$DocFileName = "技100-999.docx"
$ScriptRoot1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$ScriptRoot2 = "D:\Github\PS_Script"

# デバッグ環境に応じてパスを切り替える
if (Test-Path "D:\") {
    $ScriptRoot = $ScriptRoot2
} else {
    $ScriptRoot = $ScriptRoot1
}
$DocFilePath = Join-Path -Path $ScriptRoot -ChildPath $DocFileName

Write-Host "DocFilePath: $DocFilePath"

# 必要な型が見つからなかった場合の処理
function Process_MissingTypes {
    Write-Host "必要な型が見つかりませんでした。"
      # WdInformationの直接値を使用
      $wdVerticalPositionRelativeToPage = 1
}

# 必要な型が見つかった場合の処理
function Process_AvailableTypes {
    Write-Host "必要な型が見つかりました。処理を続行します。"

    function Find_TablesWithRoles {
        param (
            [__ComObject]$doc,
            [string[]]$roles
        )

        # #### $tablesWithRoles = @()

        
        $dummyTable = New-Object PSObject
        $tablesWithRoles = @($dummyTable)  # ダミーの初期値を追加して配列を初期化


        if ($doc.Tables -eq $null -or $doc.Tables.Count -eq 0) {
            Write-Host "ドキュメントにテーブルがありません。"
            return $tablesWithRoles
        }
        

        if ($doc.Tables -eq $null -or $doc.Tables.Count -eq 0) {
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
                    $tablesWithRoles += ,$table # 配列に追加
                }
            }
        }
        write-Host "Found Tables: $($tablesWithRoles)"
        return $tablesWithRoles
    }

    function Find_ClosestTableToTopRight {
        param (
            [__ComObject[]]$tables
        )

        $closest_Table = $null
        $closest_To_Top_Right = [double]::MaxValue

        foreach ($table in $tables) {
            $top_Position = $table.Cell(1, 1).Range.Information(1)  # Y座標
            $left_Position = $table.Cell(1, 1).Range.Information(2)  # X座標

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
        [__ComObject]$Table
        $wdInformation 

        Signature_Block([__ComObject]$doc, [string[]]$roles, [int]$wdInformation) {
            $this.Doc = $doc
            $this.Roles = $roles
            $this.wdInformation = $wdInformation

            $tablesWithRoles = Find_TablesWithRoles -doc $this.Doc -roles $this.Roles
            Write-Host "Tables With Roles: $($tablesWithRoles)"
            Write-Host "Tables With Roles Count: $($tablesWithRoles.Count)"
            if ($tablesWithRoles.Length -eq 1) {
                throw "文書内に指定された役割が含まれる表が見つかりませんでした。"
            } elseif ($tablesWithRoles.Length -eq 2) {

                # $this.Table = $tablesWithRoles[0]
                # Write-Host "First Table: $($this.Table)"

                # $this.Table = Invoke-Expression '$tablesWithRoles[0]'  # 最初のテーブルを取得
                # Write-Host "First Table: $($this.Table)"

                #$this.Table = $tablesWithRoles | Invoke-Member -MemberType Property -Name 'Item' -ArgumentList 0  # 最初のテーブルを取得
                #Write-Host "First Table: $($this.Table)"


                # $binding = "System.Reflection.BindingFlags" -as [type]

                # $this.Table = [System.__ComObject].InvokeMember("GetValue", $binding::InvokeMethod, $null, $tablesWithRoles, @(0))  # 最初のテーブルを取得
                # $this.Table = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $tablesWithRoles, 0)  # 最初のテーブルを取得
                # $this.Table = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $tablesWithRoles, @(1))  # 最初のテーブルを取得
                # $this.Table = [System.__ComObject].InvokeMember("GetValue", $binding::InvokeMethod, $null, $tablesWithRoles, @(0))  # 最初のテーブルを取得

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
        # [hashtable] Get_Signature_Coordinates([string]$cell_Set_Type) {

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

        # 役割名の位置関係からタイプを決める
            $sign_Cells = @()
            $type = "役割名上タイプ"
            for ($i = 0; $i -lt $role_Coordinates.Count - 1; $i++) {
                if ($role_Coordinates[$i].Column + 1 -ne $role_Coordinates[$i + 1].Column) {
                    $type = "役割名左タイプ"
                    break
                }
            }

        # サイン用セルの位置を求める
            foreach ($role in $role_Coordinates) {
                if ($type -eq "役割名上タイプ") {
                # 役割名上タイプ
                    $sign_Cells += [pscustomobject]@{
                        Role = $role.Role
                        Row = $role.Row + 1
                        Column = $role.Column
                    }
                } else {
                # 役割名左タイプ
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

        [array] Get_Cell_Info() {
            $cell_Info = @()
                
            for ($row = 1; $row -le 2; $row++) {
                for ($col = 1; $col -le 3; $col++) {
                    if ($row -gt $this.Table.Rows.Count -or $col -gt $this.Table.Columns.Count) {
                        throw "指定されたセルが存在しません。"
                    }
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
        
        [void] Set_Custom_Attributes_at_signature_Block() {
            $custom_Properties = $this.Doc.CustomDocumentProperties
        
            $role_To_Property_Map = @{
                "承認" = @{ Date = "承認日"; Name = "承認者" }
                "照査" = @{ Date = "照査日"; Name = "照査者" }
                "作成" = @{ Date = "作成日"; Name = "作成者" }
            }
        
    # サイン用セルの情報を変数から取得
            $sign_Cells = $this.Get_Signature_Coordinates().Sign_Cells

            foreach ($cell in $sign_Cells) {
                $role = $cell.Role
                $row = $cell.Row
                $column = $cell.Column

                $date_Property = $role_To_Property_Map[$role].Date
                $name_Property = $role_To_Property_Map[$role].Name

                if ($null -ne $custom_Properties) {
                    try {
                        $date_Value = Read_Property -wordDoc $this.Doc -PropertyName $date_Property
                    } catch {
                        $date_Value = "日付なし"
                    }

                    try {
                        $name_Value = Read_Property -wordDoc $this.Doc -PropertyName $name_Property
                    } catch {
                        $name_Value = "名前なし"
                    }
                } else {
                    $date_Value = "日付なし"
                    $name_Value = "名前なし"
                }

                $cell_Text = "$($date_Value)`n$($name_Value)"
                $this.Table.Cell($row, $column).Range.Text = $cell_Text
                $this.Table.Cell($row, $column).Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                $this.Table.Cell($row, $column).Range.Paragraphs[1].Range.Font.Size = 8
                $this.Table.Cell($row, $column).Range.Paragraphs[2].Range.Font.Size = 10
            }
        }
        
        [array] Get_Role_Above_Cell_Set() {
            $role_Above_Cell_Set = @()
                
            foreach ($role in $this.Roles) {
                $role_Index = $this.Roles.IndexOf($role) + 1
                if ($role_Index -gt $this.Table.Columns.Count) {
                    throw "指定されたセルが存在しません。"
                }
                $role_Cell = $this.Table.Cell(1, $role_Index)
                $name_Cell = $this.Table.Cell(2, $role_Index)
                    
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
                $role_Index = $this.Roles.IndexOf($role) + 1
                if ($role_Index -gt $this.Table.Columns.Count - 1) {
                    throw "指定されたセルが存在しません。"
                }
                $role_Cell = $this.Table.Cell(1, $role_Index)
                $name_Cell = $this.Table.Cell(1, $role_Index + 1)
                    
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
                $name_Cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                $name_Cell.Range.Font.Size = 8
                $name_Cell.Range.Paragraphs[2].Range.Font.Size = 10
            }
        }
    }

    # WordDocument 型のインスタンスを作成
    try {
    $wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)
    } catch {
        Write-Error "エラー: $($_)"
        exit 1
    }

    # 役割配列
    $roles = @("承認", "照査", "作成")

    # Signature_Blockクラスのインスタンスを作成
    try {
        $signature_Block = [Signature_Block]::new($wordDoc.Document, $roles, 1)  # 直接値を使用
        Write-Host "Signature_Block インスタンスが正常に作成されました。"
    } catch {
        Write-Error "エラー: $($_)"
        if ($wordDoc -ne $null) {
        $wordDoc.Close()
        }
        exit 1
    }

    # サイン欄の座標を取得
    try {
    $signature_Coordinates = $signature_Block.Get_Signature_Coordinates()
    Write-Host "サイン欄タイプ: $($signature_Coordinates.Type)"
    foreach ($sign_Cell in $signature_Coordinates.Sign_Cells) {
        Write-Host "サイン用セル：役割: $($sign_Cell.Role), 行: $($sign_Cell.Row), 列: $($sign_Cell.Column)"
        }
    } catch {
        Write-Error "エラー: $($_)"
        if ($wordDoc -ne $null) {
            $wordDoc.Close()
        }
        exit 1
    }

    # カスタム属性を設定
    try {
    $signature_Block.Set_Custom_Attributes_at_signature_Block()
    } catch {
        Write-Error "エラー: $($_)"
        if ($wordDoc -ne $null) {
            $wordDoc.Close()
        }
        exit 1
    }

    # ドキュメントを保存して閉じる
    try {
    $wordDoc.Document.Save()
    $wordDoc.Close()
    } catch {
        Write-Error "エラー: $($_)"
        if ($wordDoc -ne $null) {
            $wordDoc.Close()
        }
        exit 1
    }

    # クリーンアップ
    try {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc.Document) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc.WordApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Error "エラー: $($_)"
    }
}

# メイン処理部
$wdInformationAvailable = Test_TypeAvailability -typeName "Microsoft.Office.Interop.Word.WdInformation, Microsoft.Office.Interop.Word"

if (-not $wdInformationAvailable) {
    Process_MissingTypes
} 
Process_AvailableTypes
