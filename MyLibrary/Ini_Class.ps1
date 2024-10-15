class IniFile {
    [string]$FilePath

    IniFile([string]$filePath) {
        $this.FilePath = $filePath
    }

    [hashtable] GetContent() {
        $iniContent = @{}
        $currentSection = ""

        Write-Host "INIファイルの存在を確認中: $this.FilePath"
        if (Test-Path $this.FilePath) {
            Write-Host "INIファイルが見つかりました。内容を読み込んでいます..."
            $lines = Get-Content $this.FilePath
            foreach ($line in $lines) {
                # コメント行をスキップ
                if ($line -match "^\s*#") {
                    Write-Host "コメント行をスキップ: $line"
                    continue
                }

                # セクション名を検出
                if ($line -match "^\[(.+)\]$") {
                    $currentSection = $matches[1].Trim()
                    Write-Host "セクションを検出: $currentSection"
                    $iniContent[$currentSection] = @{}
                    continue
                }

                # 空行を検出
                if ($line -match "^\s*$") {
                    Write-Host "空行を検出: $line"
                    $currentSection = ""
                    continue
                }

                # キーと値を検出
                if ($line -match "^\s*([^=]+?)\s*=\s*(.*?)\s*$") {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    Write-Host "読み込み中: $key = $value"
                    if ($currentSection -ne "") {
                        $iniContent[$currentSection][$key] = $value
                    } else {
                        $iniContent[$key] = $value
                    }
                } else {
                    Write-Host "行が正しい形式ではありません: $line"
                }
            }
            Write-Host "INIファイルの内容を読み込みました。"
        } else {
            Write-Host "INIファイルが見つかりませんでした。"
        }
        return $iniContent
    }

    [void] SetContent([hashtable]$iniContent) {
        Write-Host "INIファイルに書き込んでいます: $this.FilePath"
        $lines = @()
        foreach ($section in $iniContent.Keys) {
            if ($iniContent[$section] -is [hashtable]) {
                $lines += "[$section]"
                foreach ($key in $iniContent[$section].Keys) {
                    $value = $iniContent[$section][$key]
                    $lines += "$key = $value"
                }
                $lines += ""
            } else {
                $lines += "$section = $($iniContent[$section])"
            }
        }
        $lines | Set-Content -Path $this.FilePath
        Write-Host "INIファイルへの書き込みが完了しました。"
    }
}