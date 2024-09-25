# PCクラスの定義
class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$WorkFolder
    [string]$CsvFolderPath
    [string]$PdfPoolFolderPath
    [string]$PdfFolderPath

    PC([string]$name, [string]$ipAddress, [string]$macAddress, [string]$workFolder) {
        $this.Name = $name
        $this.IPAddress = $ipAddress
        $this.MACAddress = $macAddress
        $this.WorkFolder = $workFolder
        $this.CsvFolderPath = Join-Path -Path $workFolder -ChildPath "#登録用csvデータ"
        $this.PdfPoolFolderPath = Join-Path -Path $workFolder -ChildPath "#登録用pdf一時保管"
        $this.PdfFolderPath = Join-Path -Path $workFolder -ChildPath "#登録用pdfデータ"

    }

    PC() {
        $this.Name = (hostname)
        
        # ネットワークインターフェース情報を取得
        $networkInterface = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
        $this.IPAddress = (Get-NetIPAddress -InterfaceIndex $networkInterface.ifIndex -AddressFamily IPv4).IPAddress
        $this.MACAddress = $networkInterface.MacAddress

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
        
        # フォルダが存在しない場合は生成
        if (-not (Test-Path -Path $this.WorkFolder)) {
            New-Item -Path $this.WorkFolder -ItemType Directory
        }
        $this.CsvFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用csvデータ"
        $this.PdfPoolFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdf一時保管"
        $this.PdfFolderPath = Join-Path -Path $this.WorkFolder -ChildPath "#登録用pdfデータ"
        
        # サブフォルダも存在しない場合は生成
        foreach ($folder in @($this.CsvFolderPath, $this.PdfPoolFolderPath, $this.PdfFolderPath)) {
            if (-not (Test-Path -Path $folder)) {
                New-Item -Path $folder -ItemType Directory
            }
        }
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

# ファイルコピー処理の実行
$fileManager.CopyFilesBasedOnCsv($pc.CsvFolderPath, $pc.WorkFolder, $settings.OutDataFolder, "C:\RealDataFolder", [ref]$successCount, [ref]$failureCount)

# 成功と失敗のカウントを表示
Write-Host "Success: $($successCount.Value)"
Write-Host "Failure: $($failureCount.Value)"
