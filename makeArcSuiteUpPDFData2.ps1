# PCクラスの定義
class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$WorkFolder
    [string]$CsvFolderPath
    [string]$PdfPoolFolderPath
    [string]$PdfFolderPath
    [hashtable]$PdfPoolHashTable

    # コンストラクタ
    PC() {
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
        $this.PdfPoolFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdf一時保管"
        $this.PdfFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdfデータ"
        $this.PdfPoolHashTable = @{}

        # フォルダの存在確認
        $this.CheckFoldersExist()

        # ハッシュテーブルの読み込み
        $this.LoadPdfPoolHashTable()
    }

    # フォルダの存在確認
    [void]CheckFoldersExist() {
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
    }

    # PDFプールフォルダのハッシュテーブルを更新
    [void]UpdatePdfPoolHashTable() {
        $this.PdfPoolHashTable.Clear()
        $files = Get-ChildItem -Path $this.PdfPoolFolderPath -Recurse -Include *.pdf, *.txt
        $totalFiles = $files.Count
        $currentFileIndex = 0

        foreach ($file in $files) {
            $currentFileIndex++
#            Write-Host "Processing file $currentFileIndex of $Files.Count : $($file.Name)`n"
            Write-Host "$currentFileIndex of $totalFiles"
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $this.PdfPoolHashTable[$file.FullName] = $hash.Hash
        }
        $this.SavePdfPoolHashTable()
    }

# ハッシュテーブルをファイルに保存
[void]SavePdfPoolHashTable() {
    if (-not (Test-Path -Path $($this.WorkFolder))) {
        New-Item -Path $($this.WorkFolder) -ItemType Directory
    }
    $json = $this.PdfPoolHashTable | ConvertTo-Json
    $json | Out-File -FilePath "$($this.WorkFolder)\PdfPoolHashTable.json" -Encoding UTF8
}

# ハッシュテーブルをファイルから読み込み
[void]LoadPdfPoolHashTable() {
    if (Test-Path -Path "$($this.WorkFolder)\PdfPoolHashTable.json") {
        $json = Get-Content -Path "$($this.WorkFolder)\PdfPoolHashTable.json" -Raw
        $this.PdfPoolHashTable = $json | ConvertFrom-Json
    } else {
        $this.PdfPoolHashTable = @{}
    }
}

    # PDFプールフォルダの状態をチェック
    [bool]HasPdfPoolFolderChanged() {
        $currentFiles = Get-ChildItem -Path $this.PdfPoolFolderPath -Recurse -Include *.pdf, *.txt
        if ($currentFiles.Count -ne $this.PdfPoolHashTable.Count) {
            return $true
        }

        foreach ($file in $currentFiles) {
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            if (-not $this.PdfPoolHashTable.ContainsKey($file.FullName) -or $this.PdfPoolHashTable[$file.FullName] -ne $hash.Hash) {
                return $true
            }
        }
        return $false
    }
}

# FileManagerクラスの定義
class FileManager {
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$pdfPoolFolderPath, [string]$pdfFolderPath, [ref]$successCount, [ref]$failureCount, [hashtable]$pdfPoolHashTable) {
        $successCount.Value = 0
        $failureCount.Value = 0
        $errorLogPath = Join-Path -Path $pdfFolderPath -ChildPath "error_log.txt"

        $csvFiles = Get-ChildItem -Path $csvFolderPath | Where-Object { $_.Name -match "_個装-???.csv" -or $_.Name -match "_図面-???.csv" -or $_.Name -match "_通知書-???.csv" }

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
                $pdfFilePath = Join-Path -Path $pdfPoolFolderPath -ChildPath "$fileName.pdf"
                $txtFilePath = Join-Path -Path $pdfPoolFolderPath -ChildPath "$fileName.txt"

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
                    }
                } else {
                    $errorMessage = "Source file not found or not in hash table: $fileName"
                    Write-Host $errorMessage
                    $errorMessage | Out-File -FilePath $errorLogPath -Append -Encoding UTF8
                    $failureCount.Value++
                }
            }
        }
    }
}

# メイン処理
$fileManager = [FileManager]::new()
$successCount = [ref]0
$failureCount = [ref]0

# PCオブジェクトの作成
$pc = [PC]::new()

# PDFプールフォルダの状態をチェックし、変化があればハッシュテーブルを更新
if ($pc.HasPdfPoolFolderChanged()) {
    $pc.UpdatePdfPoolHashTable()
}

# ファイルコピー処理の実行
$fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.PdfPoolFolderPath, $pc.PdfFolderPath, [ref]$successCount, [ref]$failureCount, $pc.PdfPoolHashTable)

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"