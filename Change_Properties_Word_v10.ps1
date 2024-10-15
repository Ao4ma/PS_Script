param (
    [string]$imagePath
)

# Set the full path for MyLibrary and join the path for PC_Class.ps1, Word_Class.ps1, and Ini_Class.ps1
$scriptFolderPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$myLibraryPath = Join-Path -Path $scriptFolderPath -ChildPath "MyLibrary"
$pcClassPath = Join-Path -Path $myLibraryPath -ChildPath "PC_Class.ps1"
$wordClassPath = Join-Path -Path $myLibraryPath -ChildPath "Word_Class.ps1"
$iniClassPath = Join-Path -Path $myLibraryPath -ChildPath "Ini_Class.ps1"

# Import the PC class, Word class, and Ini class from the respective files
. $pcClassPath
. $wordClassPath
. $iniClassPath

# INIファイルのパスを変数に格納
$iniFilePath = Join-Path -Path $scriptFolderPath -ChildPath "config_Change_Properties_Word.ini"

function ProcessDocument {
    param (
        [PC]$pc,
        [string]$filePath
    )

    if (-not (Test-Path $filePath)) {
        Write-Error "ファイルパスが無効です: $filePath"
        return
    }

    # スクリプト実行前に存在していたWordプロセスを取得
    $existingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

    # Wordアプリケーションを起動
    Write-Host "Wordアプリケーションを起動中..."
    $word = [Word]::new($filePath, $pc)

    try {
        # 文書プロパティを表示
        Write-Host "現在の文書プロパティ:"
        foreach ($property in $word.DocumentProperties) {
            Write-Host "$($property.Key): $($property.Value)"
        }

        # 承認者プロパティを設定
        # Write-Host "承認者プロパティを設定中..."
        # $wordInstanceData.SetCustomProperty("Approver", $approver)

        # 承認フラグプロパティを設定
        # $wordInstanceData.SetCustomProperty("ApprovalFlag", ($approvalFlag ? "承認" : "未承認"))

        # テーブル処理
        # $wordInstanceData.TableHandler.ProcessTable()

        # 画像処理
        # $wordInstanceData.ImageHandler.ProcessImage($imagePath)

        # カスタムオブジェクトを作成して表示
        $docProperties = $word.GetCustomObject()
        $docProperties

        # ドキュメントを保存して閉じる
        Write-Host "ドキュメントを保存して閉じています..."
        $word.Save()
        $word.Close()
    } 
    catch {
        Write-Error "エラーが発生しました: $_"
    } finally {
        # スクリプト実行後に存在するWordプロセスを取得
        $allWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue

        # スクリプト実行前に存在していたプロセスを除外して終了
        $newWordProcesses = $allWordProcesses | Where-Object { $_.Id -notin $existingWordProcesses.Id }
        foreach ($proc in $newWordProcesses) {
            Stop-Process -Id $proc.Id -Force
        }
    }

    Write-Host "カスタムプロパティが設定されました。"
}

# メイン処理

# PCクラスのインスタンスを作成し、スクリプトのあるフォルダに移動
$PcName = (hostname)
if (-not $PcName) {
    $PcName = "delld033"
}
$pc = [PC]::new($PcName, $iniFilePath)
Set-Location -Path $scriptFolderPath

# INIファイルのインスタンスを作成し、設定を読み込む
$iniFile = [IniFile]::new($iniFilePath)
$iniContent = $iniFile.GetContent()

# INIファイルから設定を読み込む
$docFileName = $iniContent["DocFile"]["DocFileName"]
$docFilePath = $iniContent["DocFile"]["DocFilePath"]

# 余分な引用符を削除
$docFilePath = $docFilePath.Trim('"')
$docFileName = $docFileName.Trim('"')

$filePath = Join-Path -Path $docFilePath -ChildPath $docFileName

# ドキュメントを処理
ProcessDocument -pc $pc -filePath $filePath

# INIファイルに設定を書き込む（必要に応じて）
$iniContent["DocFile"]["DocFileName"] = [System.IO.Path]::GetFileName($filePath)
$iniContent["DocFile"]["DocFilePath"] = [System.IO.Path]::GetDirectoryName($filePath)
$iniContent["Approver"] = $approver
$iniFile.SetContent($iniContent)