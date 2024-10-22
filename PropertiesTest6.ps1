param (
    [string]$Path,          # ドキュメントのパス
    [string]$PropertyType = "Both"  # "Builtin", "Custom", または "Both" を指定
)

# ビルトインプロパティのグループ
$BuiltinPropertiesGroup = @(
    "Title", "Subject", "Author", "Keywords", "Comments", "Template", "Last Author", 
    "Revision Number", "Application Name", "Last Print Date", "Creation Date", 
    "Last Save Time", "Total Editing Time", "Number of Pages", "Number of Words", 
    "Number of Characters", "Security", "Category", "Format", "Manager", "Company", 
    "Number of Bytes", "Number of Lines", "Number of Paragraphs", "Number of Slides", 
    "Number of Notes", "Number of Hidden Slides", "Number of Multimedia Clips", 
    "Hyperlink Base", "Number of Characters (with spaces)", "Content Type", 
    "Content Status", "Language", "Document Version"
)

# カスタムプロパティのグループ
$CustomPropertiesGroup = @("batter", "yamada", "Path")

function Get-DocumentProperties {
    param (
        [string]$Path,          # ドキュメントのパス
        [string]$PropertyType = "Both"  # "Builtin", "Custom", または "Both" を指定
    )

    try {
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Open($Path)

        $binding = "System.Reflection.BindingFlags" -as [type]
        [ref]$SaveOption = "Microsoft.Office.Interop.Word.WdSaveOptions" -as [type]

        $objHash = @{}
        $foundProperties = @()

        if ($PropertyType -eq "Builtin" -or $PropertyType -eq "Both") {
            $Properties = $doc.BuiltInDocumentProperties
            foreach ($p in $BuiltinPropertiesGroup) {
                try {
                    $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $p)
                    $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch [System.Exception] {
                    Write-Host -ForegroundColor Blue "Value not found for $p"
                }
            }
        }

        if ($PropertyType -eq "Custom" -or $PropertyType -eq "Both") {
            $Properties = $doc.CustomDocumentProperties
            foreach ($p in $CustomPropertiesGroup) {
                try {
                    $pn = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $Properties, $p)
                    $value = [System.__ComObject].InvokeMember("Value", $binding::GetProperty, $null, $pn, $null)
                    $objHash[$p] = $value
                    $foundProperties += $p
                } catch [System.Exception] {
                    Write-Host -ForegroundColor Blue "Value not found for $p"
                }
            }
        }

        # ドキュメントを保存せずに閉じる
        $doc.Close([ref]$SaveOption::wdDoNotSaveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Properties) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        Remove-Variable -Name doc, Properties

        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        Remove-Variable -Name word
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()

        Write-Host "Ready!" -ForegroundColor Green

        if ($foundProperties.Count -eq 0) {
            return 0
        } else {
            return $objHash
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        return -1
    }
}

# 関数の呼び出し例
$path = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx"
$propertyType = "Both"  # "Builtin", "Custom", または "Both"
$result = Get-DocumentProperties -Path $path -PropertyType $propertyType

if ($result -eq -1) {
    Write-Host "An error occurred during property retrieval." -ForegroundColor Red
} elseif ($result -eq 0) {
    Write-Host "No properties found." -ForegroundColor Yellow
} else {
    Write-Host "Found properties:" -ForegroundColor Green
    $result.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }
}