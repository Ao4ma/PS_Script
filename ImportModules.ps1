# ImportModules.ps1

# ここで直接パスを指定します
using module ".\MyLibrary\WordDocumentProperties.psm1"
using module ".\MyLibrary\WordDocumentUtilities.psm1"
using module ".\MyLibrary\WordDocumentSignatures.psm1"
using module ".\MyLibrary\WordDocumentChecks.psm1"
using module ".\MyLibrary\WordDocument.psm1"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# デバッグ用設定
$DocFileName = "技100-999.docx"
$ScriptRoot1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$ScriptRoot2 = "D:\Github\PS_Script"

# デバッグ環境に応じてパスを切り替える
if (Test-Path $ScriptRoot2) {
    $ScriptRoot = $ScriptRoot2
} else {
    $ScriptRoot = $ScriptRoot1
}
$DocFilePath = $ScriptRoot

# デバッグメッセージを有効にする
$DebugPreference = "Continue"

Write-Host "Creating Word Application COM object..."
# クラス外でCOMオブジェクトを作成
try {
   # $wordApp = New-Object -ComObject Word.Application
    Write-Host "Word Application COM object created successfully."
} catch {
    Write-Error "Failed to create Word Application COM object: $_"
    exit 1
}

Write-Host "Creating WordDocument instance..."
# WordDocumentクラスのインスタンスを作成
try {
    $wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)
    Write-Host "WordDocument instance created successfully."
} catch {
    Write-Error "Failed to create WordDocument instance: $_"
    exit 1
}

Write-Host "Calling Check_PC_Env..."
# メソッドの呼び出し例
try {
    $wordDoc.Check_PC_Env()
    Write-Host "Check_PC_Env completed successfully."
} catch {
    Write-Error "Check_PC_Env failed: $_"
}

Write-Host "Calling Check_Word_Library..."
try {
    $wordDoc.Check_Word_Library()
    Write-Host "Check_Word_Library completed successfully."
} catch {
    Write-Error "Check_Word_Library failed: $_"
}

Write-Host "Calling Check_Custom_Property..."
try {
    Check_Custom_Property $wordDoc
    Write-Host "Check_Custom_Property completed successfully."
} catch {
    Write-Error "Check_Custom_Property failed: $_"
}

Write-Host "Calling SetCustomPropertyAndSaveAs..."
try {
    $wordDoc.SetCustomPropertyAndSaveAs("CustomProperty31", "Value31")
    Write-Host "SetCustomPropertyAndSaveAs completed successfully."
} catch {
    Write-Error "SetCustomPropertyAndSaveAs failed: $_"
}

Write-Host "Creating WordDocument instance..."
# WordDocumentクラスのインスタンスを作成
try {
    $wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)
    Write-Host "WordDocument instance created successfully."
} catch {
    Write-Error "Failed to create WordDocument instance: $_"
    exit 1
}




Write-Host "Calling SetCustomProperty..."
try {
#   実験的にここからはクラスメソッドとした
    $wordDoc.SetCustomProperty("CustomProperty21", "Value21")
    Write-Host "SetCustomProperty completed successfully."
} catch {
    Write-Error "SetCustomProperty failed: $_"
}

<#
Write-Host "Calling SaveAs..."
try {
    $docPath = Join-Path -Path $wordDoc.DocFilePath -ChildPath $wordDoc.DocFileName
    SaveAs $docPath "$scriptRoot\temp.docx"
    Write-Host "SaveAs completed successfully."
} catch {
    Write-Error "SaveAs failed: $_"
}
#>

Write-Host "Calling SetCustomProperty..."
try {
   # SetCustomProperty
   $wordDoc.SetCustomProperty("CustomProperty1", "Value1")
    Write-Host "SetCustomProperty completed successfully."
} catch {
    Write-Error "SetCustomProperty failed: $_"
}

<#
Write-Host "Calling FillSignatures..."
try {
    # サイン欄に名前と日付を配置
    FillSignatures $wordDoc
    Write-Host "FillSignatures completed successfully."
} catch {
    Write-Error "FillSignatures failed: $_"
}
#>

Write-Host "Calling Read_Property..."
try {
    # カスタムプロパティを読み取る
    $propValue = Read_Property $wordDoc "CustomProperty21"
    Write-Host "Read Property Value: $propValue"
} catch {
    Write-Error "Read_Property failed: $_"
}

<#
Write-Host "Calling Update_Property..."
try {
    # カスタムプロパティを更新する
    Update_Property $wordDoc "CustomProperty2" "UpdatedValue"
    Write-Host "Update_Property completed successfully."
} catch {
    Write-Error "Update_Property failed: $_"
}
#>

Write-Host "Calling Delete_Property..."
try {
    # カスタムプロパティを削除する
    Delete_Property $wordDoc "CustomProperty21"
    Write-Host "Delete_Property completed successfully."
} catch {
    Write-Error "Delete_Property failed: $_"
}

Write-Host "Calling Close..."
try {
    # ドキュメントを閉じる
    $wordDoc.Close() 
    Write-Host "Close completed successfully."
} catch {
    Write-Error "Close failed: $_"
}

Write-Host "Calling Close_Word_Processes..."
try {
    # Wordプロセスを閉じる
    Close_Word_Processes $wordDoc
    Write-Host "Close_Word_Processes completed successfully."
} catch {
    Write-Error "Close_Word_Processes failed: $_"
}

Write-Host "Calling Ensure_Word_Closed..."
try {
    # Wordが閉じられていることを確認する
    Ensure_Word_Closed $wordDoc
    Write-Host "Ensure_Word_Closed completed successfully."
} catch {
    Write-Error "Ensure_Word_Closed failed: $_"
}

Write-Host "Calling WriteToFile..."
try {
    # ファイルに出力
    WriteToFile $wordDoc "$scriptRoot\output.txt" @("Line 1", "Line 2")
    Write-Host "WriteToFile completed successfully."
} catch {
    Write-Error "WriteToFile failed: $_"
}

Write-Host "Calling Get_Properties..."
try {
    # プロパティを取得する
    $properties = Get_Properties $wordDoc "Custom"
    Write-Host "Properties: $properties"
} catch {
    Write-Error "Get_Properties failed: $_"
}

Write-Host "Script completed successfully."


