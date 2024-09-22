# モジュールのインポート
Import-Module -Name "C:\\Users\\y0927\\Documents\\GitHub\\PS_Script\\ExcelProcessor.psm1"

# 新しいクラスの定義
class MainClass {
    [void] ProcessExcelFile([string]$filePath, [int]$batchSize) {
        $processor = [ExcelProcessor]::new($filePath)
        $processor.ImportExcelFile($batchSize)
    }
}

# インスタンスの作成とメソッドの呼び出し
$main = [MainClass]::new()
$main.ProcessExcelFile("S:\\技術部storage\\管理課\\PDM復旧\\ファイル1.xlsx", 1000)