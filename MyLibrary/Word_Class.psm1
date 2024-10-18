using module 'C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\PC_Class.psm1'
using module 'C:\Users\y0927\Documents\GitHub\PS_Script\MyLibrary\Ini_Class.psm1'
class Word {
    [string]$FilePath
    [PC]$PC
    [hashtable]$DocumentProperties

    Word([string]$filePath, [PC]$pc) {
        $this.FilePath = $filePath
        $this.PC = $pc
        $this.DocumentProperties = @{}

        # Wordアプリケーションを起動
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $this.WordApp = $wordApp

        # ドキュメントを開く
        $document = $wordApp.Documents.Open($filePath)
        $this.Document = $document

        # 文書プロパティを取得
        foreach ($property in $document.BuiltInDocumentProperties) {
            $this.DocumentProperties[$property.Name] = $property.Value
        }
    }

    [void] AddCustomProperty([string]$name, [string]$value) {
        $this.Document.CustomDocumentProperties.Add($name, $false, 4, $value)
    }

    [void] RemoveCustomProperty([string]$name) {
        $this.Document.CustomDocumentProperties.Item($name).Delete()
    }

    [void] RecordTableCellInfo() {
        # Implementation to record table cell info
    }

    [void] Close() {
        $this.Document.Close($false)
        $this.WordApp.Quit()
    }
}