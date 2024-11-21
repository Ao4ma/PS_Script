# ここで直接パスを指定します
using module ".\MyLibrary\WordDocumentProperties.psm1"
using module ".\MyLibrary\WordDocumentChecks.psm1"
using module ".\MyLibrary\WordDocument.psm1"
using module ".\MyLibrary\Word_Class.psm1"
using module ".\MyLibrary\Word_Table2.psm1"

param (
    [string]$docFilePath
)

# ドキュメントファイル名を変数に設定
$docFileName = "技100-999.docx"

# スクリプトのルートパスを取得
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# デバッグ用のデフォルトファイルパスを設定
$defaultDocFilePath = Join-Path -Path $scriptRoot -ChildPath $docFileName

# 作業ホームフォルダのパスを設定
$workHomeFolder = Join-Path -Path $scriptRoot -ChildPath "WorkHome"

# 作業ホームフォルダが存在しない場合は作成
if (-not (Test-Path -Path $workHomeFolder)) {
    New-Item -Path $workHomeFolder -ItemType Directory | Out-Null
}

# ログファイルとエラーファイルのパスを設定
$logFilePath = Join-Path -Path $workHomeFolder -ChildPath "log.txt"
$errorFilePath = Join-Path -Path $workHomeFolder -ChildPath "error.txt"

# ファイルパスの確認と設定
if (-not (Test-Path $docFilePath)) {
    if (Test-Path $defaultDocFilePath) {
        Write-Host "No valid file path provided. Using default file path for debugging: $defaultDocFilePath"
        $docFilePath = $defaultDocFilePath
    } else {
        $errorMessage = "No valid file path provided and default file not found."
        Write-Error $errorMessage
        $errorMessage | Out-File -FilePath $errorFilePath -Append
        exit 1
    }
}

# デバッグメッセージを有効にする
$DebugPreference = "Continue"

Write-Host "Creating Word Application COM object..."
# クラス外でCOMオブジェクトを作成
try {
    # $wordApp = New-Object -ComObject Word.Application
    Write-Host "Word Application COM object created successfully."
} catch {
    $errorMessage = "Failed to create Word Application COM object: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $errorFilePath -Append
    exit 1
}

Write-Host "Creating WordDocument instance..."
# WordDocumentクラスのインスタンスを作成
try {
    $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
    Write-Host "WordDocument instance created successfully."
} catch {
    $errorMessage = "Failed to create WordDocument instance: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $errorFilePath -Append
    exit 1
}

Write-Host "Calling Check_PC_Env..."
# メソッドの呼び出し例
try {
    $wordDoc.Check_PC_Env()
    Write-Host "Check_PC_Env completed successfully."
} catch {
    $errorMessage = "Check_PC_Env failed: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $errorFilePath -Append
}

Write-Host "Calling Check_Word_Library..."
try {
    $wordDoc.Check_Word_Library()
    Write-Host "Check_Word_Library completed successfully."
} catch {
    $errorMessage = "Check_Word_Library failed: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $errorFilePath -Append
}

Write-Host "Calling checkCustomProperty..."
try {
    $wordDoc.checkCustomProperty2()
    Write-Host "checkCustomProperty completed successfully."
} catch {
    $errorMessage = "checkCustomProperty failed: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $errorFilePath -Append
}

# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property2("CustomProperty1")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("CustomProperty2")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("CustomProperty21")
Write-Host "Read Property Value: $propValue"

# カスタムプロパティを更新する
# $wordDoc.Update_Property("CustomProperty2", "UpdatedValue")

# カスタムプロパティを削除する
# $wordDoc.Delete_Property("CustomProperty21")
$wordDoc.CheckCustomProperty2()
# $wordDoc.Delete_Property("CustomProperty1")
# $wordDoc.Check_Custom_Property()

# ドキュメントを閉じる
$wordDoc.Close()

# ログファイルに成功メッセージを記録
"Script completed successfully." | Out-File -FilePath $logFilePath -Append