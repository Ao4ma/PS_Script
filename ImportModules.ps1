# ImportModules.ps1
using module "$PSScriptRoot\MyLibrary\WordDocument.psm1"
using module "$PSScriptRoot\MyLibrary\WordDocumentProperties.psm1"
using module "$PSScriptRoot\MyLibrary\WordDocumentUtilities.psm1"
using module "$PSScriptRoot\MyLibrary\WordDocumentSignatures.psm1"
using module "$PSScriptRoot\MyLibrary\WordDocumentChecks.psm1"

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

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)

# メソッドの呼び出し例
$wordDoc.Check_PC_Env()
$wordDoc.Check_Word_Library()
$wordDoc.Check_Custom_Property()
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty21", "Value21", "C:\Users\y0927\Documents\GitHub\PS_Script\sample_temp.docx")

# サイン欄に名前と日付を配置
$wordDoc.FillSignatures()

# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property("CustomProperty2")
Write-Host "Read Property Value: $propValue"

# カスタムプロパティを更新する
$wordDoc.Update_Property("CustomProperty2", "UpdatedValue")

# カスタムプロパティを削除する
$wordDoc.Delete_Property("CustomProperty2")

# ドキュメントを閉じる
$wordDoc.Close()

# Wordプロセスを閉じる
$wordDoc.Close_Word_Processes()

# Wordが閉じられていることを確認する
$wordDoc.Ensure_Word_Closed()

# ファイルに出力
$wordDoc.WriteToFile("C:\path\to\output.txt", @("Line 1", "Line 2"))

# プロパティを取得する
$properties = $wordDoc.Get_Properties("Custom")
Write-Host "Properties: $properties"