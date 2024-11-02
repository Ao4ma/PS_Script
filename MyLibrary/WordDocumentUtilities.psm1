# MyLibrary/WordDocumentUtilities.psm1
# Nullチェックメソッド
[bool] CheckNull([object]$obj, [string]$message) {
    if ($null -eq $obj) {
        Write-Host $message -ForegroundColor Red
        return $true
    }
    return $false
}

# ドキュメントを別名で保存するメソッド
[void] SaveAs([string]$NewFilePath) {
    $this.Document.SaveAs([ref]$NewFilePath)
}

# ドキュメントを閉じるメソッド
[void] Close() {
    $this.Document.Close([ref]$false)  # false を指定して変更を保存しない
    $this.WordApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ファイルに内容を書き込むメソッド
[void] WriteToFile([string]$FilePath, [array]$Content) {
    if ($Content.Count -eq 0) {
        Write-Host "No content found. Deleting previous output file if it exists."
        if (Test-Path $FilePath) {
            Remove-Item $FilePath
        }
    } else {
        $Content | Out-File -FilePath $FilePath
    }
}

# Wordプロセスを閉じるメソッド
[void] Close_Word_Processes() {
    Write-Host "IN: Close_Word_Processes"
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($existingWordProcesses) {
        foreach ($process in $existingWordProcesses) {
            Stop-Process -Id $process.Id -Force
        }
    }
    Write-Host "OUT: Close_Word_Processes"
}

# Wordが閉じられていることを確認するメソッド
[void] Ensure_Word_Closed() {
    Write-Host "IN: Ensure_Word_Closed"
    $newWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($newWordProcesses) {
        foreach ($process in $newWordProcesses) {
            Stop-Process -Id $process.Id -Force
        }
    }
    Write-Host "OUT: Ensure_Word_Closed"
}