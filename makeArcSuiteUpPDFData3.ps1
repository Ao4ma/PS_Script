class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$WorkFolder
    [string]$CsvFolderPath
    [string]$PdfPoolFolderPath
    [string]$PdfFolderPath
    [hashtable]$PdfPoolHashTable
    [hashtable]$FileNameToFullPathHashTable

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
        $this.FileNameToFullPathHashTable = @{}

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
        $this.FileNameToFullPathHashTable.Clear()
        $files = Get-ChildItem -Path $this.PdfPoolFolderPath -Recurse -Include *.pdf, *.txt
        $totalFiles = $files.Count
        $currentFileIndex = 0

        foreach ($file in $files) {
            $currentFileIndex++
            Write-Host "$currentFileIndex of $totalFiles"
            $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
            $this.PdfPoolHashTable[$file.FullName] = $hash.Hash
            $this.FileNameToFullPathHashTable[$file.Name] = $file.FullName
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

        $jsonFileNameToFullPath = $this.FileNameToFullPathHashTable | ConvertTo-Json
        $jsonFileNameToFullPath | Out-File -FilePath "$($this.WorkFolder)\FileNameToFullPathHashTable.json" -Encoding UTF8
    }

    # ハッシュテーブルをファイルから読み込み
    [void]LoadPdfPoolHashTable() {
        if (Test-Path -Path "$($this.WorkFolder)\PdfPoolHashTable.json") {
            $json = Get-Content -Path "$($this.WorkFolder)\PdfPoolHashTable.json" -Raw
            $this.PdfPoolHashTable = $json | ConvertFrom-Json
        } else {
            $this.PdfPoolHashTable = @{}
        }

        if (Test-Path -Path "$($this.WorkFolder)\FileNameToFullPathHashTable.json") {
            $jsonFileNameToFullPath = Get-Content -Path "$($this.WorkFolder)\FileNameToFullPathHashTable.json" -Raw
            $this.FileNameToFullPathHashTable = $jsonFileNameToFullPath | ConvertFrom-Json
        } else {
            $this.FileNameToFullPathHashTable = @{}
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