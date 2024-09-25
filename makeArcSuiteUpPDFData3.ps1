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
    [hashtable]$FilePathHashTable

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
        $this.FilePathHashTable = @{}

        # フォルダの存在確認
        $this.CheckFoldersExist()

        # ハッシュテーブルとファイルパスハッシュテーブルの読み込み
        $this.LoadHashTables()
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

    # PDFプールフォルダのハッシュテーブルとファイルパスハッシュテーブルを更新
    [void]UpdateHashTables() {
        $this.PdfPoolHashTable.Clear()
        $this.FilePathHashTable.Clear()
        $files = Get-ChildItem -Path $this.PdfPoolFolderPath -Recurse -Include *.pdf, *.txt
        $totalFiles = $files.Count
        $currentFileIndex = 0

        foreach ($file in $files) {
            $currentFileIndex++
            Write-Host "$currentFileIndex of $totalFiles"
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $this.PdfPoolHashTable[$file.FullName] = $hash.Hash
            $this.FilePathHashTable[$file.BaseName] = $file.FullName
        }
        $this.SaveHashTables()
    }

    # ハッシュテーブルとファイルパスハッシュテーブルをファイルに保存
    [void]SaveHashTables() {
        if (-not (Test-Path -Path $($this.WorkFolder))) {
            New-Item -Path $($this.WorkFolder) -ItemType Directory
        }
        $jsonHash = $this.PdfPoolHashTable | ConvertTo-Json
        $jsonHash | Out-File -FilePath "$($this.WorkFolder)\PdfPoolHashTable.json" -Encoding UTF8

        $jsonFilePath = $this.FilePathHashTable | ConvertTo-Json
        $jsonFilePath | Out-File -FilePath "$($this.WorkFolder)\FilePathHashTable.json" -Encoding UTF8
    }

    # ハッシュテーブルとファイルパスハッシュテーブルをファイルから読み込み
    [void]LoadHashTables() {
        if (Test-Path -Path "$($this.WorkFolder)\PdfPoolHashTable.json") {
            $jsonHash = Get-Content -Path "$($this.WorkFolder)\PdfPoolHashTable.json" -Raw
            $this.PdfPoolHashTable = $jsonHash | ConvertFrom-Json
        } else {
            $this.PdfPoolHashTable = @{}
        }

        if (Test-Path -Path "$($this.WorkFolder)\FilePathHashTable.json") {
            $jsonFilePath = Get-Content -Path "$($this.WorkFolder)\FilePathHashTable.json" -Raw
            $this.FilePathHashTable = $jsonFilePath | ConvertFrom-Json
        } else {
            $this.FilePathHashTable = @{}
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
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$pdfPoolFolderPath, [string]$pdfFolderPath, [ref]$successCount, [ref]$failureCount, [hashtable]$filePathHashTable) {
        $successCount.Value = 0
        $failureCount.Value = 0
        $errorLogPath = Join-Path -Path $pdfFolderPath -ChildPath "error_log.txt"

        $csvFiles = Get-ChildItem -Path $csvFolderPath | Where-Object { $_.Name -like "_個装-???.csv" -or $_.Name -like "_図面-???.csv" -or $_.Name -like "_通知書-???.csv" }
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

                if ($filePathHashTable.ContainsKey($fileName)) {
                    $sourceFilePath = $filePathHashTable[$fileName]
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
                    $errorMessage = "Source file not found in hash table: $fileName"
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

# PDFプールフォルダの状態をチェックし、変化があればハッシュテーブルとファイルパスハッシュテーブルを更新
if ($pc.HasPdfPoolFolderChanged()) {
    $pc.UpdateHashTables()
}

# ファイルコピー処理の実行
$fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.PdfPoolFolderPath, $pc.PdfFolderPath, [ref]$successCount, [ref]$failureCount, $pc.FilePathHashTable)

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"