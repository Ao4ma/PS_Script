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
                $this.WorkFolder = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)"  # デフォルトの作業フォルダ
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
        $this.PdfFilePathMap = @{}  # 新しい連想配列の初期化

        # フォルダの存在確認
        $this.CheckFoldersExist()

        # ハッシュテーブルと連想配列の読み込み
        $this.LoadHashTable("PdfPoolHashTable.json", [ref]$this.PdfPoolHashTable)
        $this.LoadHashTable("FilePathHashTable.json", [ref]$this.FilePathHashTable)
        $this.LoadHashTable("PdfFilePathMap.json", [ref]$this.PdfFilePathMap)
        Write-Host "Exiting PC constructor"
    }

    # フォルダの存在確認
    [void]CheckFoldersExist() {
        Write-Host "Entering CheckFoldersExist"
        # フォルダが存在しない場合はエラーをスロー
        if (-not (Test-Path -Path $this.WorkFolder)) {
            throw "Work folder does not exist: $($this.WorkFolder)"
        }
        
        # サブフォルダも存在しない場合はエラーをスロー
        foreach ($folder in @($this.CsvFolderPath, $this.PdfPoolFolderPath, $this.PdfFolderPath)) {
            if (-not (Test-Path -Path $folder)) {
                throw "Required folder does not exist: $folder"
            }
        }
        Write-Host "Exiting CheckFoldersExist"
    }

    # ハッシュテーブルと連想配列を更新
    [void]UpdateHashTable([string]$folderPath, [string]$fileExtensions, [ref]$hashTable) {
        Write-Host "Entering UpdateHashTable"
        $hashTable.Value.Clear()
        $this.PdfFilePathMap.Clear()  # 新しい連想配列のクリア
        $extensions = $fileExtensions -split ","
        $files = Get-ChildItem -Path $folderPath -Recurse | Where-Object { $extensions -contains $_.Extension }
        $totalFiles = $files.Count
        $currentFileIndex = 0

        foreach ($file in $files) {
            $currentFileIndex++
            Write-Host "Processing file $currentFileIndex of $($totalFiles): $($file.FullName)"
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $hashTable.Value[$file.FullName] = $hash.Hash
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $this.PdfFilePathMap[$fileName] = $file.FullName  # 新しい連想配列に追加
        }
        $this.SaveHashTable("PdfPoolHashTable.json", [ref]$this.PdfPoolHashTable)
        $this.SaveHashTable("PdfFilePathMap.json", [ref]$this.PdfFilePathMap)
        Write-Host "Exiting UpdateHashTable"
    }

    # ハッシュテーブルと連想配列をファイルに保存
    [void]SaveHashTable([string]$fileName, [ref]$hashTable) {
        Write-Host "Entering SaveHashTable"
        if (-not (Test-Path -Path $this.WorkFolder)) {
            New-Item -Path $this.WorkFolder -ItemType Directory -Force
        }
        $json = $hashTable.Value | ConvertTo-Json
        $filePath = Join-Path -Path $this.WorkFolder -ChildPath $fileName
        Write-Host "Saving hash table to: $filePath"
        $json | Out-File -FilePath $filePath -Encoding UTF8
        Write-Host "Exiting SaveHashTable"
    }

    # ハッシュテーブルと連想配列をファイルから読み込み
    [void]LoadHashTable([string]$fileName, [ref]$hashTable) {
        Write-Host "Entering LoadHashTable"
        $filePath = Join-Path -Path $this.WorkFolder -ChildPath $fileName
        if (Test-Path -Path $filePath) {
            # JSON ファイルの読み込み
            $json = Get-Content -Path $filePath -Raw

            # JSON を PSCustomObject に変換
            $psCustomObject = $json | ConvertFrom-Json

            # PSCustomObject を Hashtable に変換
            $hashTable.Value = @{}
            foreach ($key in $psCustomObject.PSObject.Properties.Name) {
                $hashTable.Value[$key] = $psCustomObject.$key
            }
        } else {
            $hashTable.Value = @{}
        }
        Write-Host "Exiting LoadHashTable"
    }

    # フォルダの状態をチェック
    [bool]HasFolderChanged([string]$folderPath, [string]$fileExtensions, [hashtable]$hashTable) {
        Write-Host "Entering HasFolderChanged"
        $extensions = $fileExtensions -split ","
        $currentFiles = Get-ChildItem -Path $folderPath -Recurse | Where-Object { $extensions -contains $_.Extension }
        if ($currentFiles.Count -ne $hashTable.Count) {
            Write-Host "Exiting HasFolderChanged with result: $true"
            return $true
        }

        foreach ($file in $currentFiles) {
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


# FileManagerクラスの定義
class FileManager {
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$pdfPoolFolderPath, [string]$pdfFolderPath, [ref]$successCount, [ref]$failureCount, [hashtable]$pdfPoolHashTable) {
        Write-Host "Entering CopyFilesBasedOnCsv"
        $successCount.Value = 0
        $failureCount.Value = 0
        $errorLogPath = Join-Path -Path $pdfFolderPath -ChildPath "error_log.txt"

        $csvFiles = Get-ChildItem -Path $csvFolderPath | Where-Object { 
            $_.Name -match "_(個装|図面|通知書)-\d{3}\.csv" 
        }

        foreach ($csvFile in $csvFiles) {
            $csvData = Import-Csv -Path $csvFile.FullName
            $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
            $csvDestinationFolder = Join-Path -Path $pdfFolderPath -ChildPath $csvFileName

            # CSVファイル名のフォルダを作成
            if (-not (Test-Path -Path $csvDestinationFolder)) {
                New-Item -Path $csvDestinationFolder -ItemType Directory
            }

            foreach ($row in $csvData) {
                $fileName = $row.'関連付け用ファイル名'
                $subFolderPath = Join-Path -Path $pdfPoolFolderPath -ChildPath $fileName
                $pdfFilePath = Join-Path -Path $subFolderPath -ChildPath "$fileName.pdf"
                $txtFilePath = Join-Path -Path $subFolderPath -ChildPath "$fileName.txt"

                $sourceFilePath = if (Test-Path $pdfFilePath) { $pdfFilePath } elseif (Test-Path $txtFilePath) { $txtFilePath } else { $null }

                if ($sourceFilePath -and $pdfPoolHashTable.ContainsKey($sourceFilePath)) {
                    $destinationFilePath = Join-Path -Path $csvDestinationFolder -ChildPath (Get-Item $sourceFilePath).Name

                    Write-Host "Copying file: $sourceFilePath to $destinationFilePath"

                    try {
                        Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -ErrorAction Stop
                        $successCount.Value++
                        Write-Host "Successfully copied: $sourceFilePath"
                    } catch {
                        $errorMessage = "Failed to copy $sourceFilePath to $destinationFilePath"
                        Write-Host $errorMessage
                        $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                        $failureCount.Value++
                        throw $errorMessage  # エラーが発生した場合にスクリプトを停止
                    }
                } else {
                    $errorMessage = "Source file not found or not in hash table: $fileName"
                    Write-Host $errorMessage
                    $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                    $failureCount.Value++
                    throw $errorMessage  # エラーが発生した場合にスクリプトを停止
                }
            }
        }
        Write-Host "Exiting CopyFilesBasedOnCsv"
    }
}

# メイン処理
Write-Host "Starting main script"
$fileManager = [FileManager]::new()
$successCount = [ref]0
$failureCount = [ref]0

# PCオブジェクトの作成
$pc = [PC]::new()

# PDFプールフォルダの状態をチェックし、変化があればハッシュテーブルを更新
if ($pc.HasFolderChanged($pc.PdfPoolFolderPath, "*.pdf, *.txt", $pc.PdfPoolHashTable)) {
    $pc.UpdateHashTable($pc.PdfPoolFolderPath, "*.pdf, *.txt", [ref]$pc.PdfPoolHashTable)
}

# ファイルパスの状態をチェックし、変化があればハッシュテーブルを更新
if ($pc.HasFolderChanged($pc.CsvFolderPath, "*.csv", $pc.FilePathHashTable)) {
    $pc.UpdateHashTable($pc.CsvFolderPath, "*.csv", [ref]$pc.FilePathHashTable)
}

# ファイルコピー処理の実行
try {
    $fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.PdfPoolFolderPath, $pc.PdfFolderPath, [ref]$successCount, [ref]$failureCount, $pc.PdfPoolHashTable)
} catch {
    Write-Host "An error occurred: $_"
    break  # エラーが発生した場合にスクリプトを停止
}

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"
Write-Host "Ending main script"
