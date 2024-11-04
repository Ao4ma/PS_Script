# Microsoft.Office.Interop.Wordアセンブリをロード
Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

# WordアプリケーションのCOMオブジェクトを作成
$word = New-Object -ComObject Word.Application
if ($null -eq $word) {
    Write-Error "Failed to create Word Application COM object."
    exit
}

# DisplayAlertsを無効に設定
$word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone

# ドキュメントのパス
$docPath = "D:\Github\PS_Script\技100-999.docx"

# ドキュメントを開く
try {
    $document = $word.Documents.Open($docPath)
    if ($null -eq $document) {
        Write-Error "Failed to open document: $docPath"
        $word.Quit()
        exit
    }
    Write-Output "Document opened successfully."
} catch {
    Write-Error "Failed to open document: $_"
    $word.Quit()
    exit
}

# 必要な操作をここに追加

# ドキュメントを保存して閉じる
$document.Save()
$document.Close()
$word.Quit()