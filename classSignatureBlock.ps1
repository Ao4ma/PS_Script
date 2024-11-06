class SignatureBlock {
    $Doc
    [string[]]$Roles
    $Table

    SignatureBlock($doc, [string[]]$roles) {
        $this.Doc = $doc
        $this.Roles = $roles
        $this.Table = $this.FindSignatureTable()
        if ($null -eq $this.Table) {
            throw "文書フォーマットが違います。サイン欄の条件に合う表が見つかりませんでした。"
        }
        $this.ValidateRoles()
    }

    [void] ValidateRoles() {
        $expectedRoles = @("承認", "照査", "作成")
        $tableRoles = @()
        for ($col = 1; $col -le 3; $col++) {
            $roleCell = $this.Table.Cell(1, $col)
            $roleText = $roleCell.Range.Text.Trim()
            $tableRoles += $roleText
        }

        if ($expectedRoles -ne $tableRoles) {
            throw "サイン欄フォーマットが違います。役割の文字列が一致しません。"
        }
    }

    [hashtable] Get-SignatureCoordinates([string]$cellSetType) {
        if ($cellSetType -eq "above") {
            $signatureOrigin = $this.Table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            $signatureDiagonal = $this.Table.Cell(2, 3).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
        } elseif ($cellSetType -eq "left") {
            $signatureOrigin = $this.Table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            $signatureDiagonal = $this.Table.Cell(1, 3).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
        } else {
            throw "Invalid cell set type. Use 'above' or 'left'."
        }
        
        return @{
            Origin = $signatureOrigin
            Diagonal = $signatureDiagonal
        }
    }

    [array] Get-CellInfo() {
        $cellInfo = @()
        
        for ($row = 1; $row -le 2; $row++) {
            for ($col = 1; $col -le 3; $col++) {
                $cell = $this.Table.Cell($row, $col)
                $cellRange = $cell.Range
                $cellText = $cellRange.Text.Trim()
                $cellPosition = $cellRange.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
                
                $cellInfo += [pscustomobject]@{
                    CellNumber = "R$row-C$col"
                    Text = $cellText
                    Position = $cellPosition
                }
            }
        }
        
        return $cellInfo
    }

    [void] Set-CustomAttributes() {
        $customProperties = $this.Doc.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]

        $roleToPropertyMap = @{
            "承認" = @{ Date = "承認日"; Name = "承認者" }
            "照査" = @{ Date = "照査日"; Name = "照査者" }
            "作成" = @{ Date = "作成日"; Name = "作成者" }
        }

        foreach ($role in $this.Roles) {
            $roleIndex = $this.Roles.IndexOf($role) + 1
            $nameCell = $this.Table.Cell(2, $roleIndex)

            $dateProperty = $roleToPropertyMap[$role].Date
            $nameProperty = $roleToPropertyMap[$role].Name

            $dateValue = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $dateProperty).Value
            $nameValue = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $nameProperty).Value

            $nameCell.Range.Text = "$dateValue`n$nameValue"
            $nameCell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
            $nameCell.Range.Paragraphs[1].Range.Font.Size = 8
            $nameCell.Range.Paragraphs[2].Range.Font.Size = 10
        }
    }

    [array] Get-RoleAboveCellSet() {
        $roleAboveCellSet = @()
        
        foreach ($role in $this.Roles) {
            $roleCell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 1))
            $nameCell = $this.Table.Cell(2, ($this.Roles.IndexOf($role) + 1))
            
            $roleAboveCellSet += [pscustomobject]@{
                Role = $role
                RoleCell = $roleCell
                NameCell = $nameCell
            }
        }
        
        return $roleAboveCellSet
    }

    [array] Get-RoleLeftCellSet() {
        $roleLeftCellSet = @()
        
        foreach ($role in $this.Roles) {
            $roleCell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 1))
            $nameCell = $this.Table.Cell(1, ($this.Roles.IndexOf($role) + 2))
            
            $roleLeftCellSet += [pscustomobject]@{
                Role = $role
                RoleCell = $roleCell
                NameCell = $nameCell
            }
        }
        
        return $roleLeftCellSet
    }

    [void] Set-RoleNameCells([string]$cellSetType) {
        if ($cellSetType -eq "above") {
            $roleCellSet = $this.Get-RoleAboveCellSet()
        } elseif ($cellSetType -eq "left") {
            $roleCellSet = $this.Get-RoleLeftCellSet()
        } else {
            throw "Invalid cell set type. Use 'above' or 'left'."
        }

        foreach ($cellSet in $roleCellSet) {
            $role = $cellSet.Role
            $nameCell = $cellSet.NameCell
            $date = Get-Date -Format "yyyy/MM/dd"
            $name = "名前"
            
            $nameCell.Range.Text = "$date`n$name"
            $nameCell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
            $nameCell.Range.Font.Size = 8
            $nameCell.Range.Paragraphs[2].Range.Font.Size = 10
        }
    }

    [void] FindSignatureTable() {
        $highestTable = $null
        $highestPosition = [double]::MinValue

        foreach ($table in $this.Doc.Tables) {
            $position = $table.Cell(1, 1).Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
            if ($position -gt $highestPosition) {
                $highestPosition = $position
                $highestTable = $table
            }
        }

        return $highestTable
    }
}

# Wordアプリケーションを起動
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# ドキュメントを開く
$doc = $word.Documents.Open("C:\path\to\your\document.docx")

# 役割配列
$roles = @("承認", "照査", "作成")

# SignatureBlockクラスのインスタンスを作成
try {
    $signatureBlock = [SignatureBlock]::new($doc, $roles)
} catch {
    Write-Error "エラー: $_"
    $doc.Close()
    $word.Quit()
    exit 1
}

# サイン欄の座標を取得
$signatureCoordinates = $signatureBlock.Get-SignatureCoordinates("left")
Write-Output "サイン欄原点の座標: $($signatureCoordinates.Origin)"
Write-Output "サイン欄対角の座標: $($signatureCoordinates.Diagonal)"

# サイン欄内のセル情報を取得
$cellInfo = $signatureBlock.Get-CellInfo()
$cellInfo | ForEach-Object { Write-Output "セル番号: $($_.CellNumber), 位置: $($_.Position), 座標: $($_.Position)" }

# 役割上セルセットの情報を取得
$roleAboveCellSet = $signatureBlock.Get-RoleAboveCellSet()
$roleAboveCellSet | ForEach-Object { Write-Output "役割: $($_.Role), 役割セル: $($_.RoleCell), 名前セル: $($_.NameCell)" }

# 役割左セルセットの情報を取得
$roleLeftCellSet = $signatureBlock.Get-RoleLeftCellSet()
$roleLeftCellSet | ForEach-Object { Write-Output "役割: $($_.Role), 役割セル: $($_.RoleCell), 名前セル: $($_.NameCell)" }

# カスタム属性を設定
$signatureBlock.Set-CustomAttributes()

# 役割名セルを設定
$signatureBlock.Set-RoleNameCells("left")

# ドキュメントを保存して閉じる
$doc.Save()
$doc.Close()
$word.Quit()