# MyLibrary/WordDocumentUtilities.psm1
function CheckNull {
    param ($this, $obj, $message)
    if ($null -eq $obj) {
        Write-Host $message -ForegroundColor Red
        return $true
    }
    return $false
}

function SaveAs {
    param ($this, $NewFilePath)
    $this.Document.SaveAs([ref]$NewFilePath)
}

function Close {
    param ($this)
    $this.Document.Close([ref]$false)  # false を指定して変更を保存しない
    $this.WordApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Document) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function WriteToFile {
    param ($this, $FilePath, $Content)
    if ($Content.Count -eq 0) {
        Write-Host "No content found. Deleting previous output file if it exists."
        if (Test-Path $FilePath) {
            Remove-Item $FilePath
        }
    } else {
        $Content | Out-File -FilePath $FilePath
    }
}

function Close_Word_Processes {
    param ($this)
    Write-Host "IN: Close_Word_Processes"
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($existingWordProcesses) {
        foreach ($process in $existingWordProcesses) {
            Stop-Process -Id $process.Id -Force
        }
    }
    Write-Host "OUT: Close_Word_Processes"
}

function Ensure_Word_Closed {
    param ($this)
    Write-Host "IN: Ensure_Word_Closed"
    $newWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($newWordProcesses) {
        foreach ($process in $newWordProcesses) {
            Stop-Process -Id $process.Id -Force
        }
    }
    Write-Host "OUT: Ensure_Word_Closed"
}