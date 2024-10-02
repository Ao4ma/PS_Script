# 検索対象のディレクトリパス
$targetDirectories = @(
    "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#eValue元データ"
#    "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdf_生成場所"
)

# 検索する文字列のリスト
$searchStrings = @(
    "000637da853782C2",
    "000400ca880406D2-1",
    "0003eedc880406C2",
    "0006e1c8880519E2",
    "0003fb59880519C2",
    "00040050880519D2",
    "0003eedd880520C2",
    "000400cb880520C2-1",
    "00087498取説No2762"
)

# コピー先のディレクトリパス
$destinationDirectory = "S:\技術部storage\管理課\管理課共有資料\ArcSuite\#eValue-AS移行データ(本番)\#登録用pdf_生成場所\temp"

foreach ($targetDirectory in $targetDirectories) {
    Write-Host "Searching in directory: $targetDirectory"
    
    # 指定されたディレクトリ内のすべてのファイルを再帰的に取得
    $files = Get-ChildItem -Path $targetDirectory -Recurse

    foreach ($file in $files) {
        foreach ($searchString in $searchStrings) {
            # ファイル名に検索文字列が含まれているか確認
            if ($file.Name -like "*$searchString*") {
                # ファイルのフルパスを表示
                Write-Host "Found file: $($file.FullName)"
                
                # ファイルをコピー
                Copy-Item -Path $file.FullName -Destination $destinationDirectory -Force
                Write-Host "Copied file to: $destinationDirectory"
            }
        }
    }
}