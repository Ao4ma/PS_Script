# Microsoft.Office.Interop.Word アセンブリを読み込む
$assemblyPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"
try {
    Add-Type -Path $assemblyPath -ErrorAction Stop
    Write-Output "アセンブリが $assemblyPath から正常に読み込まれました"
} catch {
    Write-Error "アセンブリを $assemblyPath から読み込めませんでした。エラー: $_"
    exit 1
}

# アセンブリ内のすべての型をリストアップする
$types = [System.Reflection.Assembly]::LoadFrom($assemblyPath).GetTypes()
$wordTypes = $types | Where-Object { $_.Namespace -eq "Microsoft.Office.Interop.Word" }
$wordTypes | ForEach-Object { Write-Output $_.FullName }
