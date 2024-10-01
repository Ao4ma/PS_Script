class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$WorkFolder
    [string]$CsvFolderPath
    [string]$PdfPoolFolderPath
    [string]$PdfFolderPath
    [hashtable]$PdfPoolHashTable
    [hashtable]$FilePathHashTable
    [hashtable]$PdfFilePathMap  # 新しい連想配列

    # コンストラクタ
    PC() {
        Write-Host "Entering PC constructor"
        $this.Name = (hostname)
        
        # ネットワークインターフェース情報を取得
        $networkInterface = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
        $this.IPAddress = (Get-NetIPAddress -InterfaceIndex $networkInterface.ifIndex -AddressFamily IPv4).IPAddress
        $this.MACAddress = $networkInterface.MacAddress

        # PC名に基づいて作業フォルダを設定
        switch ($this.Name) {
            "delld033" {
                $this.WorkFolder = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)_test"  # デフォルトの作業フォルダ
            }
            "AsusTuf" {
                $this.WorkFolder = "D:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)"  # デフォルトの作業フォルダ
            }
            default {
                $this.WorkFolder = "C:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)"  # デフォルトの作業フォルダ
            }
        }
        
        $this.CsvFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用csvデータ"
        $this.PdfPoolFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdf_生成場所"
        $this.PdfFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdfデータ"
        $this.PdfPoolHashTable = @{}
        $this.FilePathHashTable = @{}
        $this.PdfFilePathMap = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)  # 大文字小文字を無視する連想配列の初期化
    }

    # ハッシュテーブルと連想配列を更新
    [void]UpdateHashTable([string]$folderPath, [string]$fileExtensions, [ref]$hashTable, [int]$batchSize = 1000) {
        Write-Host "Entering UpdateHashTable"
        
        # 確認メッセージを表示
        $confirmation = Read-Host "Do you want to initialize the hash table? (yes/no) [default: yes]"
        if ([string]::IsNullOrWhiteSpace($confirmation) -or $confirmation -eq "yes") {
            $confirmation = "yes"
        }

        if ($confirmation -ne "yes") {
            Write-Host "Hash table initialization canceled."
            Write-Host "Exiting UpdateHashTable"
            return
        }

        $hashTable.Value.Clear()
        if ($folderPath -eq $this.PdfPoolFolderPath) {
            $this.PdfFilePathMap.Clear()  # 新しい連想配列のクリア
        }
        $extensions = $fileExtensions -split "," | ForEach-Object { "*$($_.TrimStart('*'))" }
        $files = Get-ChildItem -Path $folderPath -Recurse -File | 
            Where-Object { 
                ($ext = $_.Extension); 
                ($extensions | ForEach-Object { $ext -like $_ }) -contains $true 
            }
        $totalFiles = $files.Count
        $currentFileIndex = 0

        foreach ($file in $files) {
            $currentFileIndex++
            Write-Host "Processing file $currentFileIndex of $($totalFiles): $($file.FullName)"
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $hashTable.Value[$file.FullName] = $hash.Hash
            if ($folderPath -eq $this.PdfPoolFolderPath) {
                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $this.PdfFilePathMap[$fileName] = $file.FullName  # 新しい連想配列に追加
            }

            # バッチサイズごとに処理を一時停止
            if ($currentFileIndex % $batchSize -eq 0) {
                Write-Host "Processed $currentFileIndex files. Pausing for a moment..."
                Start-Sleep -Seconds 2
            }
        }

        # ハッシュテーブルを保存
        if ($folderPath -eq $this.PdfPoolFolderPath) {
            $this.SaveHashTable("PdfPoolHashTable.json", [ref]$this.PdfPoolHashTable)
            $this.SaveHashTable("PdfFilePathMap.json", [ref]$this.PdfFilePathMap)
        } elseif ($folderPath -eq $this.CsvFolderPath) {
            $this.SaveHashTable("FilePathHashTable.json", [ref]$this.FilePathHashTable)
        }
        
        Write-Host "Exiting UpdateHashTable"
    }

    # ハッシュテーブルを保存
    [void]SaveHashTable([string]$fileName, [ref]$hashTable) {
        Write-Host "Entering SaveHashTable"
        $json = $hashTable.Value | ConvertTo-Json -Compress
        Set-Content -Path $fileName -Value $json -Encoding UTF8
        Write-Host "Exiting SaveHashTable"
    }

    # ハッシュテーブルを読み込み
    [void]LoadHashTable([string]$fileName, [ref]$hashTable, [int]$batchSize = 1000) {
        Write-Host "Entering LoadHashTable"
        if (Test-Path -Path $fileName) {
            $json = Get-Content -Path $fileName -Encoding UTF8 | Out-String
            $hashTable.Value = $json | ConvertFrom-Json
        }
        Write-Host "Exiting LoadHashTable"
    }

    # フォルダの変更をチェック
    [bool]HasFolderChanged([string]$folderPath, [string]$fileExtensions, [hashtable]$hashTable) {
        Write-Host "Entering HasFolderChanged"
        $extensions = $fileExtensions -split "," | ForEach-Object { "*$($_.TrimStart('*'))" }
        $files = Get-ChildItem -Path $folderPath -Recurse -File | 
            Where-Object { 
                ($ext = $_.Extension); 
                ($extensions | ForEach-Object { $ext -like $_ }) -contains $true 
            }

        foreach ($file in $files) {
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            if (-not $hashTable.ContainsKey($file.FullName) -or $hashTable[$file.FullName] -ne $hash.Hash) {
                Write-Host "Exiting HasFolderChanged with result: $true"
                return $true
            }
        }
        Write-Host "Exiting HasFolderChanged with result: $false"
        return $false
    }
}

class FileManager {
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$pdfPoolFolderPath, [string]$pdfFolderPath, [ref]$successCount, [ref]$failureCount, [hashtable]$pdfPoolHashTable, [hashtable]$pdfFilePathMap) {
        Write-Host "Entering CopyFilesBasedOnCsv"
        $successCount.Value = 0
        $failureCount.Value = 0
        $errorLogPath = Join-Path -Path $pdfFolderPath -ChildPath "error_log.txt"
        $obsoleteLogPath = Join-Path -Path $pdfFolderPath -ChildPath "廃図作成処理ログ.txt"

        $csvFiles = Get-ChildItem -Path $csvFolderPath | Where-Object { 
            $_.Name -match "_(個装|図面|通知書).*-\d{3}\.csv" 
        }

        foreach ($csvFile in $csvFiles) {
            $csvData = Import-Csv -Path $csvFile.FullName
            $csvFileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
            $csvFileFolder = Join-Path -Path $pdfFolderPath -ChildPath $csvFileNameWithoutExtension

            # CSVファイル名のフォルダを作成
            if (-not (Test-Path -Path $csvFileFolder)) {
                New-Item -Path $csvFileFolder -ItemType Directory -Force
            }

            foreach ($row in $csvData) {
                $fileName = $row.'関連付け用ファイル名'.Trim()
                
                # pdfFilePathMap からフルパスを取得
                if ($pdfFilePathMap.ContainsKey($fileName)) {
                    $pdfFilePath = $pdfFilePathMap[$fileName]
                } else {
                    $pdfFilePath = $null
                    $errorMessage = "Source file not found in PdfFilePathMap: $fileName"
                    Write-Host $errorMessage
                    $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                    $failureCount.Value++
                    continue
                }
            
                $txtFilePath = if ($pdfFilePath) { [System.IO.Path]::ChangeExtension($pdfFilePath, ".txt") } else { $null }
            
                $sourceFilePath = if ($pdfFilePath -and (Test-Path $pdfFilePath)) { 
                    $pdfFilePath 
                } elseif ($txtFilePath -and (Test-Path $txtFilePath)) { 
                    $txtFilePath 
                } else { 
                    $null 
                }
            
                if ($sourceFilePath -and $pdfPoolHashTable.ContainsKey($sourceFilePath)) {
                    $destinationFilePath = Join-Path -Path $csvFileFolder -ChildPath (Get-Item $sourceFilePath).Name
            
                    Write-Host "Copying file: $sourceFilePath to $destinationFilePath"
            
                    try {
                        Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -ErrorAction Stop
                        $successCount.Value++
                        Write-Host "Successfully copied: $sourceFilePath"
                    } catch {
                        $errorMessage = "Failed to copy $sourceFilePath to $destinationFilePath. Error: $_"
                        Write-Host $errorMessage
                        $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                        $failureCount.Value++
                    }
                } else {
                    $errorMessage = "Source file not found or not in hash table: $fileName"
                    Write-Host $errorMessage
                    $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                    $failureCount.Value++
                }

                # 行データに「廃」が含まれている場合の処理
                if ($row -like "*廃*") {
                    $txtFileName = "$fileName.txt"
                    if ($pdfFilePathMap.ContainsKey($fileName)) {
                        $txtFilePath = $pdfFilePathMap[$fileName]
                        if (Test-Path $txtFilePath) {
                            $destinationTxtFilePath = Join-Path -Path $csvFileFolder -ChildPath (Get-Item $txtFilePath).Name
                            Write-Host "Copying file: $txtFilePath to $destinationTxtFilePath"
                            try {
                                Copy-Item -Path $txtFilePath -Destination $destinationTxtFilePath -ErrorAction Stop
                                Write-Host "Successfully copied: $txtFilePath"
                            } catch {
                                $errorMessage = "Failed to copy $txtFilePath to $destinationTxtFilePath. Error: $_"
                                Write-Host $errorMessage
                                $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                            }
                        } else {
                            $errorMessage = "TXT file not found: $txtFilePath"
                            Write-Host $errorMessage
                            $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                        }
                    } else {
                        $errorMessage = "TXT file not found in PdfFilePathMap: $txtFileName"
                        Write-Host $errorMessage
                        $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                    }
                }
            }
        }
        
        Write-Host "Exiting CopyFilesBasedOnCsv"
    }

    [void]VerifyFilesInFolders([string]$pdfFolderPath) {
        Write-Host "Entering VerifyFilesInFolders"
        $folders = Get-ChildItem -Path $pdfFolderPath -Directory
        $discrepancyLogPath = Join-Path -Path $pdfFolderPath -ChildPath "discrepancy_log.txt"
        $obsoleteLogPath = Join-Path -Path $pdfFolderPath -ChildPath "廃図作成処理ログ.txt"

        foreach ($folder in $folders) {
            $csvFileName = "$($folder.Name).csv"
            $csvFilePath = Join-Path -Path $pdfFolderPath -ChildPath $csvFileName

            if (Test-Path -Path $csvFilePath) {
                $csvData = Import-Csv -Path $csvFilePath
                $discrepancies = @()

                foreach ($row in $csvData) {
                    $fileName = $row.'関連付け用ファイル名'.Trim()
                    $filePath = Join-Path -Path $folder.FullName -ChildPath "$fileName.pdf"

                    if (-not (Test-Path -Path $filePath)) {
                        $discrepancies += $fileName
                        $obsoleteFilePath = Join-Path -Path $folder.FullName -ChildPath ("$fileName" + "廃図.txt")
                        "Obsolete drawing" | Out-File -FilePath $obsoleteFilePath -Encoding UTF8
                        "Created obsolete file: $obsoleteFilePath" | Out-File -FilePath $obsoleteLogPath -Append -Encoding UTF8
                    }
                }

                if ($discrepancies.Count -gt 0) {
                    "Discrepancies found in folder: $($folder.Name)" | Out-File -FilePath $discrepancyLogPath -Append -Encoding UTF8
                    $discrepancies | ForEach-Object { "Missing file: $_" | Out-File -FilePath $discrepancyLogPath -Append -Encoding UTF8 }
                } else {
                    "No discrepancies found in folder: $($folder.Name)" | Out-File -FilePath $discrepancyLogPath -Append -Encoding UTF8
                }
            } else {
                "CSV file not found for folder: $($folder.Name)" | Out-File -FilePath $discrepancyLogPath -Append -Encoding UTF8
            }
        }

        Write-Host "Exiting VerifyFilesInFolders"
    }
}

# メイン処理
Write-Host "Starting main script"
$fileManager = [FileManager]::new()
$successCount = [ref]0
$failureCount = [ref]0

# PCオブジェクトの作成
$pc = [PC]::new()

# エラーログのパスを定義
$errorLogPath = Join-Path -Path $pc.PdfFolderPath -ChildPath "error_log.txt"

# PDFプールフォルダの状態をチェックし、変化があればハッシュテーブルと連想配列を更新
if ($pc.HasFolderChanged($pc.PdfPoolFolderPath, "*.pdf, *.txt", $pc.PdfPoolHashTable)) {
    $pc.UpdateHashTable($pc.PdfPoolFolderPath, "*.pdf, *.txt", [ref]$pc.PdfPoolHashTable, 1000)
} else {
    $pc.LoadHashTable("PdfFilePathMap.json", [ref]$pc.PdfFilePathMap, 1000)
}

# ファイルパスの状態をチェックし、変化があればハッシュテーブルを更新
if ($pc.HasFolderChanged($pc.CsvFolderPath, "*.csv", $pc.FilePathHashTable)) {
    $pc.UpdateHashTable($pc.CsvFolderPath, "*.csv", [ref]$pc.FilePathHashTable, 1000)
}

# ファイルコピー処理の実行
try {
    $fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.PdfPoolFolderPath, $pc.PdfFolderPath, [ref]$successCount, [ref]$failureCount, $pc.PdfPoolHashTable, $pc.PdfFilePathMap)
} catch {
    Write-Host "An error occurred: $_"
    $errorMessage = "An error occurred during file copy. Error: $_"
    $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
}

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"

# フォルダ内のファイルを確認
$fileManager.VerifyFilesInFolders($pc.PdfFolderPath)

Write-Host "Ending main script"
