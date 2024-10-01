# エラーログファイルのパスを設定
$errorLogPath = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)_test\error_log.txt"

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
        
        # 確認メッセージを表示し、5秒間待機を削除
        Write-Host "Initializing the hash table..."
    
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
                Write-Host "Folder has changed: $($file.FullName)"
                Write-Host "Exiting HasFolderChanged"
                return $true
            }
        }
        Write-Host "No changes detected in folder: $folderPath"
        Write-Host "Exiting HasFolderChanged"
        return $false
    }

    # ハッシュテーブルを保存
    [void]SaveHashTable([string]$filePath, [ref]$hashTable) {
        try {
            Write-Host "Saving hash table to $filePath"
            $json = $hashTable.Value | ConvertTo-Json -Depth 10
            $json | Out-File -FilePath $filePath -Encoding UTF8
            Write-Host "Successfully saved hash table to $filePath"
        } catch {
            $errorMessage = "Error saving hash table to $filePath. Error: $_"
            $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
            Write-Host $errorMessage
        }
    }

    # ハッシュテーブルを読み込む
    [void]LoadHashTable([string]$filePath, [ref]$hashTable) {
        Write-Host "Entering LoadHashTable"
        if (Test-Path -Path $filePath) {
            $json = Get-Content -Path $filePath -Raw
            $hashTable.Value = $json | ConvertFrom-Json
        }
        Write-Host "Exiting LoadHashTable"
    }
}


class FileManager {
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$pdfPoolFolderPath, [string]$pdfFolderPath, [ref]$successCount, [ref]$failureCount, [hashtable]$pdfPoolHashTable, [hashtable]$pdfFilePathMap) {
        Write-Host "Entering CopyFilesBasedOnCsv"
        $successCount.Value = 0
        $failureCount.Value = 0

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

                # ファイル名の最後に半角スペースが入っている場合に対応
                $fileName = $fileName.TrimEnd()

                # pdfFilePathMap からフルパスを取得
                if ($pdfFilePathMap.ContainsKey($fileName)) {
                    $sourceFilePath = $pdfFilePathMap[$fileName]
                    $destinationFilePath = Join-Path -Path $csvFileFolder -ChildPath "$fileName.pdf"
                    Write-Host "Copying file: $sourceFilePath to $destinationFilePath"
                    try {
                        Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force
                        $successCount.Value++
                    } catch {
                        Write-Host "Failed to copy file: $fileName"
                        $failureCount.Value++
                    }
                } else {
                    Write-Host "Source file not found in PdfFilePathMap: $fileName"
                    $failureCount.Value++
                }
            }
        }

        Write-Host "Exiting CopyFilesBasedOnCsv"
    }

    [void]VerifyFilesInFolders([string]$pdfFolderPath, [string]$logFolderPath) {
        Write-Host "Entering VerifyFilesInFolders"

        if (-not (Test-Path -Path $pdfFolderPath)) {
            Write-Host "Path not found: $pdfFolderPath"
            Write-Host "Exiting VerifyFilesInFolders"
            return
        }

        $folders = Get-ChildItem -Path $pdfFolderPath -Directory

        foreach ($folder in $folders) {
            $csvFileName = "$($folder.Name).csv"
            $csvFilePath = Join-Path -Path $pdfFolderPath -ChildPath $csvFileName
            $folderLogPath = Join-Path -Path $logFolderPath -ChildPath "$($folder.Name).log"

            if (Test-Path -Path $csvFilePath) {
                $csvData = Import-Csv -Path $csvFilePath
                $totalFiles = $csvData.Count
                $successPdfCount = 0
                $successTxtCount = 0
                $successObsoletePdfCount = 0
                $successObsoleteTxtCount = 0
                $generatedObsoleteTxtCount = 0
                $generatedObsoleteFiles = @()

                foreach ($row in $csvData) {
                    $fileName = $row.'関連付け用ファイル名'.Trim()
                    $fileNameTrimmed = $fileName.TrimEnd()  # ファイル名の最後に半角スペースが入っている場合に対応
                    $pdfFilePath = Join-Path -Path $folder.FullName -ChildPath "$fileNameTrimmed.pdf"
                    $txtFilePath = Join-Path -Path $folder.FullName -ChildPath "$fileNameTrimmed.txt"

                    if (Test-Path -Path $pdfFilePath) {
                        if ($fileName -like "*廃*") {
                            $successObsoletePdfCount++
                        } else {
                            $successPdfCount++
                        }
                    }

                    if (Test-Path -Path $txtFilePath) {
                        if ($fileName -like "*廃*") {
                            $successObsoleteTxtCount++
                        } else {
                            $successTxtCount++
                        }
                    } elseif ($fileName -like "*廃*") {
                        $generatedObsoleteTxtCount++
                        $obsoleteFilePath = Join-Path -Path $folder.FullName -ChildPath "$fileNameTrimmed-Obsolete.txt"
                        "Obsolete drawing" | Out-File -FilePath $obsoleteFilePath -Encoding UTF8
                        "Created obsolete file: $obsoleteFilePath" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                        $generatedObsoleteFiles += $obsoleteFilePath
                    }

                    # トリム処理したファイル名とフルパスをログに記載
                    if ($fileName -ne $fileNameTrimmed) {
                        "Trimmed file name: $fileNameTrimmed (Original: $fileName)" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                        "Full path: $pdfFilePath" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                    }
                }

                $totalSuccessFiles = $successPdfCount + $successTxtCount + $successObsoletePdfCount + $successObsoleteTxtCount
                $discrepancyCount = $totalFiles - $totalSuccessFiles

                "Total files: $totalFiles" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Successfully copied PDFs (without '廃'): $successPdfCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Successfully copied TXTs (without '廃'): $successTxtCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Successfully copied PDFs (with '廃'): $successObsoletePdfCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Successfully copied TXTs (with '廃'): $successObsoleteTxtCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Generated obsolete TXTs: $generatedObsoleteTxtCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                "Discrepancy count: $discrepancyCount" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8

                if ($generatedObsoleteFiles.Count -gt 0) {
                    "Generated obsolete files:" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                    $generatedObsoleteFiles | ForEach-Object { $_ | Out-File -FilePath $folderLogPath -Append -Encoding UTF8 }
                }

                if ($discrepancyCount -gt 0) {
                    "Discrepancies found in folder: $($folder.Name)" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                    $csvData | Where-Object { -not (Test-Path -Path (Join-Path -Path $folder.FullName -ChildPath "$($_.'関連付け用ファイル名'.TrimEnd()).pdf")) } | ForEach-Object {
                        "Missing file: $($_.'関連付け用ファイル名'.TrimEnd())" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                    }
                } else {
                    "No discrepancies found in folder: $($folder.Name)" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
                }
            } else {
                "CSV file not found: $csvFileName" | Out-File -FilePath $folderLogPath -Append -Encoding UTF8
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

# ログフォルダのパスを定義
$logFolderPath = Join-Path -Path $pc.WorkFolder -ChildPath "logs"
if (-not (Test-Path -Path $logFolderPath)) {
    New-Item -Path $logFolderPath -ItemType Directory -Force
} else {
    # ログフォルダ内のログファイルをクリア
    Get-ChildItem -Path $logFolderPath -Filter *.txt | ForEach-Object { Clear-Content -Path $_.FullName }
}

# PDFプールフォルダの状態をチェックし、変化があればハッシュテーブルと連想配列を更新
try {
    if ($pc.HasFolderChanged($pc.PdfPoolFolderPath, "*.pdf, *.txt", $pc.PdfPoolHashTable)) {
        $pc.UpdateHashTable($pc.PdfPoolFolderPath, "*.pdf, *.txt", [ref]$pc.PdfPoolHashTable, 1000)
    } else {
        $pc.LoadHashTable("PdfFilePathMap.json", [ref]$pc.PdfFilePathMap)
    }
} catch {
    $errorMessage = "Error checking or updating PDF pool folder. Error: $_"
    $errorMessage | Out-File -FilePath $logFolderPath\error_log.txt -Append -Encoding UTF8
    Write-Host $errorMessage
}

# Csvファイルの状態をチェックし、変化があればハッシュテーブルを更新
try {
    if ($pc.HasFolderChanged($pc.CsvFolderPath, "*.csv", $pc.FilePathHashTable)) {
        $pc.UpdateHashTable($pc.CsvFolderPath, "*.csv", [ref]$pc.FilePathHashTable, 1000)
    }
} catch {
    $errorMessage = "Error checking or updating CSV folder. Error: $_"
    $errorMessage | Out-File -FilePath $logFolderPath\error_log.txt -Append -Encoding UTF8
    Write-Host $errorMessage
}

# ファイルコピー処理の実行
try {
    $csvFolders = Get-ChildItem -Path $pc.CsvFolderPath -Directory
    foreach ($csvFolder in $csvFolders) {
        $folderErrorLogPath = Join-Path -Path $logFolderPath -ChildPath "$($csvFolder.Name)_error_log.txt"
        try {
            Write-Host "Processing CSV folder: $($csvFolder.FullName)"
            $fileManager.CopyFilesBasedOnCsv($csvFolder.FullName, $pc.PdfPoolFolderPath, $pc.PdfFolderPath, [ref]$successCount, [ref]$failureCount, $pc.PdfPoolHashTable, $pc.PdfFilePathMap)
        } catch {
            $errorMessage = "An error occurred: $_"
            $errorMessage | Out-File -FilePath $folderErrorLogPath -Append -Encoding UTF8
            Write-Host $errorMessage
        }
    }
} catch {
    $errorMessage = "An error occurred during file copy. Error: $_"
    $errorMessage | Out-File -FilePath $logFolderPath\error_log.txt -Append -Encoding UTF8
    Write-Host $errorMessage
}

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"

# フォルダ内のファイルを確認
try {
    $fileManager.VerifyFilesInFolders($pc.PdfFolderPath, $logFolderPath)
} catch {
    $errorMessage = "An error occurred during file verification. Error: $_"
    $errorMessage | Out-File -FilePath $logFolderPath\error_log.txt -Append -Encoding UTF8
    Write-Host $errorMessage
}

Write-Host "Ending main script"