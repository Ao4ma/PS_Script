# ImportModules4.ps1

# ここで直接パスを指定します
using module ".\MyLibrary\WordDocumentProperties.psm1"
using module ".\MyLibrary\WordDocumentChecks.psm1"
using module ".\MyLibrary\WordDocument.psm1"
using module ".\MyLibrary\Word_Class.psm1"
using module ".\MyLibrary\Word_Table2.psm1"

param (
    [string]$docFilePath
)

# スクリプトのファイル名を取得
$scriptFileName = [System.IO.Path]::GetFileName($MyInvocation.MyCommand.Definition)

# スクリプトのファイル名を表示
Write-Host "Running script: $scriptFileName"

if (-not $docFilePath) {
    Write-Error "Usage: .\$scriptFileName -docFilePath <path to doc file>"
    exit 1
}

# スクリプトのルートパスを取得
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# 作業ホームフォルダのパスを設定
$workHomeFolder = Join-Path -Path $scriptRoot -ChildPath "WorkHome"

# 作業ホームフォルダが存在しない場合は作成
if (-not (Test-Path -Path $workHomeFolder)) {
    New-Item -Path $workHomeFolder -ItemType Directory | Out-Null
}

# 作業ホームフォルダに移動
Set-Location -Path $workHomeFolder

# フルパスを取得
$fullDocFilePath = Resolve-Path -Path $docFilePath

# WordDocumentクラスのインスタンスを作成
$wordDoc = [WordDocument]::new($fullDocFilePath, $scriptRoot)

# メソッドの呼び出し例
$wordDoc.Check_PC_Env()
$wordDoc.Check_Word_Library()
$wordDoc.SetCustomProperty("CustomProperty1", "Value1")
$wordDoc.Check_Custom_Property()
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty21", "Value21")

# 新しいインスタンスを作成してから操作を続行
$wordDoc = [WordDocument]::new($fullDocFilePath, $scriptRoot)

# サイン欄に名前と日付を配置
# $wordDoc.FillSignatures()

# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property("CustomProperty1")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property("CustomProperty2")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property("CustomProperty21")
Write-Host "Read Property Value: $propValue"

# カスタムプロパティを更新する
# $wordDoc.Update_Property("CustomProperty2", "UpdatedValue")

# カスタムプロパティを削除する
$wordDoc.Delete_Property("CustomProperty21")
$wordDoc.Check_Custom_Property()
# $wordDoc.Delete_Property("CustomProperty1")
# $wordDoc.Check_Custom_Property()

# ドキュメントを閉じる
$wordDoc.Close()