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
        foreach ($file in $files) {
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $this.PdfPoolHashTable[$file.FullName] = $hash.Hash
        }
        $this.SavePdfPoolHashTable()
    }

    # ハッシュテーブルをファイルに保存
    [void]SavePdfPoolHashTable() {
        $json = $this.PdfPoolHashTable | ConvertTo-Json
        $json | Out-File -FilePath "$this.WorkFolder\PdfPoolHashTable.json" -Encoding UTF8
    }

    # ハッシュテーブルをファイルから読み込み
    [void]LoadPdfPoolHashTable() {
        if (Test-Path -Path "$this.WorkFolder\PdfPoolHashTable.json") {
            $json = Get-Content -Path "$this.WorkFolder\PdfPoolHashTable.json" -Raw
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
    [void]CopyFilesBasedOnCsv([string]$csvFolderPath, [string]$sourceFolder, [string]$destinationFolder, [string]$realDataFolder, [ref]$successCount, [ref]$failureCount) {
        $successCount.Value = 0
        $failureCount.Value = 0

        $csvFiles = Get-ChildItem -Path $csvFolderPath -Filter "*_個装*-???.csv", "*_図面*-???.csv", "*_通知書*-???.csv"

        foreach ($csvFile in $csvFiles) {
            $csvData = Import-Csv -Path $csvFile.FullName

            foreach ($row in $csvData) {
                $sourceFilePath = Join-Path -Path $sourceFolder -ChildPath $row.FileName
                $destinationFilePath = Join-Path -Path $destinationFolder -ChildPath $row.FileName

                if (Test-Path $sourceFilePath) {
                    try {
                        Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force
                        $successCount.Value++
                    } catch {
                        Write-Host "Failed to copy $sourceFilePath to $destinationFilePath"
                        $failureCount.Value++
                    }
                } else {
                    Write-Host "Source file not found: $sourceFilePath"
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
$fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.WorkFolder, "C:\OutDataFolder", "C:\RealDataFolder", [ref]$successCount, [ref]$failureCount)

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"