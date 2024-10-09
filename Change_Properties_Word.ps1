function Import-InteropAssembly {
    Param (
        [string]$assemblyName = "Microsoft.Office.Interop.Word"
    )

    $scriptPath = $MyInvocation.MyCommand.Path
    $iniFilePath = [System.IO.Path]::ChangeExtension($scriptPath, ".ini")

    # INIファイルからアセンブリパスを読み込む
    if (Test-Path $iniFilePath) {
        $assemblyPath = Get-Content $iniFilePath
    } else {
        $assemblyPath = $null
    }

    switch ($true) {
        ($assemblyPath -and (Test-Path $assemblyPath)) {
            Write-Output "$assemblyName is found in INI file. Using the existing assembly."
            Add-Type -Path $assemblyPath
        }
        ($null -eq $assemblyPath -or -not (Test-Path $assemblyPath)) {
            Write-Output "$assemblyName is not found in INI file or path does not exist. Searching in Windows directory..."

            # Windowsディレクトリ下をサーチ
            $assemblyPath = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" | Select-Object -First 1 -ExpandProperty FullName

            if ($assemblyPath) {
                Write-Output "$assemblyName is found in Windows directory. Using the existing assembly."
                Set-Content -Path $iniFilePath -Value $assemblyPath
                Add-Type -Path $assemblyPath
            } else {
                Write-Output "$assemblyName is not found in Windows directory. Installing from NuGet..."

                # NuGetプロバイダーのインストール
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
                    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
                }

                # Microsoft.Office.Interop.Wordのインストール
                if (-not (Get-Package -Name Microsoft.Office.Interop.Word -ErrorAction SilentlyContinue)) {
                    Install-Package -Name Microsoft.Office.Interop.Word -Source nuget.org -Scope CurrentUser
                }

                # パッケージのインポート
                Add-Type -AssemblyName "Microsoft.Office.Interop.Word"
            }
        }
    }

    # Microsoft.Office.Interop.Word アセンブリのパスを手動で指定
    # $assemblyPath = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll"

    if (-Not (Test-Path $assemblyPath)) {
        Write-Host "Assembly path does not exist: $assemblyPath" -Foreground Red
        exit
    }
    Add-Type -Path $assemblyPath
}

# GACをチェックしてMicrosoft.Office.Interop.Wordが存在するか確認
Import-InteropAssembly

function Get-ConfigValue {
    Param (
        [string]$key,
        [string]$iniFilePath
    )

    if (Test-Path $iniFilePath) {
        $config = Get-Content $iniFilePath | ForEach-Object {
            $line = $_.Trim()
            if ($line -match "^\s*([^=]+)\s*=\s*(.+)\s*$") {
                [PSCustomObject]@{ Key = $matches[1]; Value = $matches[2] }
            }
        }

        $entry = $config | Where-Object { $_.Key -eq $key }
        if ($entry) {
            return $entry.Value
        }
    }

    return $null
}

function Set-ConfigValue {
    Param (
        [string]$key,
        [string]$value,
        [string]$iniFilePath
    )

    if (Test-Path $iniFilePath) {
        $config = Get-Content $iniFilePath
        $updated = $false

        $config = $config | ForEach-Object {
            $line = $_.Trim()
            if ($line -match "^\s*([^=]+)\s*=\s*(.+)\s*$" -and $matches[1] -eq $key) {
                $updated = $true
                "$key=$value"
            } else {
                $_
            }
        }

        if (-not $updated) {
            $config += "$key=$value"
        }

        Set-Content -Path $iniFilePath -Value $config
    } else {
        Set-Content -Path $iniFilePath -Value "$key=$value"
    }
}

function Install-InteropAssembly {
    Param (
        [string]$assemblyName = "Microsoft.Office.Interop.Word"
    )

    Write-Output "$assemblyName is not found. Installing from NuGet..."

    # NuGetプロバイダーのインストール
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    }

    # Microsoft.Office.Interop.Wordのインストール
    if (-not (Get-Package -Name Microsoft.Office.Interop.Word -ErrorAction SilentlyContinue)) {
        Install-Package -Name Microsoft.Office.Interop.Word -Source nuget.org -Scope CurrentUser
    }

    # パッケージのインポート
    Add-Type -AssemblyName "Microsoft.Office.Interop.Word"
}

# INIファイルのパスを取得
$scriptPath = $MyInvocation.MyCommand.Path
$iniFilePath = [System.IO.Path]::ChangeExtension($scriptPath, ".ini")

# アセンブリパスを取得
$assemblyPath = Get-ConfigValue -key "AssemblyPath" -iniFilePath $iniFilePath

if ($assemblyPath -and (Test-Path $assemblyPath)) {
    Write-Output "Using assembly from INI file."
    Add-Type -Path $assemblyPath
} else {
    Write-Output "Searching for assembly in Windows directory..."
    $assemblyPath = Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" | Select-Object -First 1 -ExpandProperty FullName

    if ($assemblyPath) {
        Write-Output "Assembly found in Windows directory."
        Set-ConfigValue -key "AssemblyPath" -value $assemblyPath -iniFilePath $iniFilePath
        Add-Type -Path $assemblyPath
    } else {
        Install-InteropAssembly
        $assemblyPath = (Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Recurse -Filter "Microsoft.Office.Interop.Word.dll" | Select-Object -First 1 -ExpandProperty FullName)
        Set-ConfigValue -key "AssemblyPath" -value $assemblyPath -iniFilePath $iniFilePath
    }
}

# 印章IMGのパスを取得
$imagePath = Get-ConfigValue -key "ImagePath" -iniFilePath $iniFilePath
if (-not $imagePath) {
    $imagePath = "C:\Users\y0927\Documents\GitHub\PS_Script\社長印.tif"
    Set-ConfigValue -key "ImagePath" -value $imagePath -iniFilePath $iniFilePath
}

function Get-OfficeDocBuiltInProperties {
    [OutputType([Hashtable])]
    Param
    (
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    $result = @{}
    $binding = "System.Reflection.BindingFlags" -as [type]
    $properties = $Document.BuiltInDocumentProperties
    
    foreach($property in $properties)
    {
        $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
        trap [system.exception]
        {
            continue
        }
        $result.Add($pn, [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null))
    }
 
    return $result
}
 
function Get-OfficeDocBuiltInProperty {
    [OutputType([string],$null)]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try {
        $comObject = $Document.BuiltInDocumentProperties($PropertyName)
        $binding = "System.Reflection.BindingFlags" -as [type]
        $val = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$comObject,$null)
        return $val
    } catch {
        return $null
    }
}
 
function Set-OfficeDocBuiltInProperty {
    [OutputType([boolean])]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [string] $Value,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try {
        $comObject = $Document.BuiltInDocumentProperties($PropertyName)
        $binding = "System.Reflection.BindingFlags" -as [type]
        [System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,$Value)
        return $true
    } catch {
        return $false
    }
}
 
function Get-OfficeDocCustomProperties {
    [OutputType([HashTable])]
    Param
    (
         [Parameter(Mandatory=$true, Position=2)]
         [System.__ComObject] $Document
    )
 
    $result = @{}
    $binding = "System.Reflection.BindingFlags" -as [type]
    $properties = $Document.CustomDocumentProperties
    foreach($property in $properties)
    {
        $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
        trap [system.exception]
        {
            continue
        }
        $result.Add($pn, [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null))
    }
 
    return $result
}
 
function Get-OfficeDocCustomProperty {
    [OutputType([string], $null)]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try {
        $comObject = $Document.CustomDocumentProperties($PropertyName)
        $binding = "System.Reflection.BindingFlags" -as [type]
        $val = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$comObject,$null)
        return $val
    } catch {
        return $null
    }
}
 
function Set-OfficeDocCustomProperty {
    [OutputType([boolean])]
    Param
    (
         [Parameter(Mandatory=$true)]
         [string] $PropertyName,
         [Parameter(Mandatory=$true)]
         [string] $Value,
         [Parameter(Mandatory=$true)]
         [System.__ComObject] $Document
    )
 
    try
    {
        $customProperties = $Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName,$false, 4, $Value
        try
        {
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod,$null,$customProperties,$arrayArgs) | out-null
        } 
        catch [system.exception] 
        {
            $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
            [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
            [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
        }
        return $true
    } 
    catch
    {
        return $false
    }
}

# デバッグ情報の追加
Write-Host "Start Word and load a document..." -Foreground Yellow
$app = New-Object -ComObject Word.Application
$app.visible = $false
Write-Host "Word application started."

$docPath = "C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx"
if (Test-Path $docPath) {
    Write-Host "Document path exists: $docPath"
    $doc = $app.Documents.Open($docPath, $false, $false, $false)
    Write-Host "Document opened: $docPath"
} else {
    Write-Host "Document path does not exist: $docPath" -Foreground Red
    exit
}

# 1つ目のテーブルを取得
$table = $doc.Tables.Item(1)
Write-Host "First table retrieved."

# テーブルのプロパティを取得
$rows = $table.Rows.Count
$columns = $table.Columns.Count
Write-Host "Table properties retrieved: Rows=$rows, Columns=$columns"

# 各セルの情報を取得
foreach ($row in 1..$rows) {
    foreach ($col in 1..$columns) {
        try {
            $cell = $table.Cell($row, $col)
            $cellText = $cell.Range.Text
            Write-Host "Row: $row, Column: $col, Text: $cellText"
        } catch {
            Write-Host "Row: $row, Column: $col, Text: (cell not found)" -Foreground Red
        }
    }
}

# 1つ目のセルを取得
$cell = $table.Cell(2, 6)
Write-Host "Cell (2, 6) retrieved."

# セルの座標とサイズを取得
$left = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdHorizontalPositionRelativeToPage)
$top = $cell.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdVerticalPositionRelativeToPage)
$width = $cell.Width
$height = $cell.Height
Write-Host "Cell coordinates and size retrieved: Left=$left, Top=$top, Width=$width, Height=$height"

# 画像のサイズを設定
$imageWidth = 50
$imageHeight = 50

# 画像の中央位置を計算
$imageLeft = $left + ($width - $imageWidth) / 2
$imageTop = $top + ($height - $imageHeight) / 2

# 既存の画像を削除（もしあれば）
foreach ($shape in $doc.Shapes) {
    if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdInlineShapeType]::wdInlineShapePicture) {
        $shape.Delete()
    }
}
Write-Host "Existing images deleted."

# 新しい画像を挿入
$shape = $doc.Shapes.AddPicture("C:\Users\y0927\Documents\GitHub\PS_Script\社長印.tif", $false, $true, $imageLeft, $imageTop, $imageWidth, $imageHeight)
Write-Host "New image inserted."

# 画像のプロパティを変更
$shape.LockAspectRatio = $false
$shape.Width = 50
$shape.Height = 50
Write-Host "Image properties modified."

write-host "`nAll BUILT IN Properties:" -Foreground Yellow
Get-OfficeDocBuiltInProperties $doc
 
write-host "`nWrite to BUILT IN author property:" -Foreground Yellow
$result = Set-OfficeDocBuiltInProperty "Author" "Mr. Robot" $doc
write-host "Result: $result"
 
write-host "`nRead BUILT IN author again:" -Foreground Yellow
Get-OfficeDocBuiltInProperty "Author" $doc
 
write-host "`nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc
 
write-host "`nWrite a CUSTOM property:" -Foreground Yellow
$result = Set-OfficeDocCustomProperty "Hacked by" "fsociety" $doc
write-host "Result: $result"
 
write-host "`nRead back the CUSTOM property:" -Foreground Yellow
Get-OfficeDocCustomProperty "Hacked by" $doc

write-host "`2 nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc

write-host "`nWrite a CUSTOM property:" -Foreground Yellow
$result = Set-OfficeDocCustomProperty "Batter" "Otani" $doc
write-host "Result: $result"
 
write-host "`2 nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc

write-host "`nAll CUSTOM Properties (none if new document):" -Foreground Yellow
Get-OfficeDocCustomProperties $doc

write-host "`nSave document and close Word..." -Foreground Yellow

# ファイルが読み取り専用でないことを確認
if (-not (Test-Path $docPath -PathType Leaf -ErrorAction SilentlyContinue)) {
    Write-Host "File is read-only: $docPath" -Foreground Red
    exit
}

# ドキュメントを保存して閉じる
$doc.Save()
$doc.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null

# アプリケーションを終了する
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null

# ガベージコレクションを実行して、未解放のCOMオブジェクトを解放する
[gc]::collect()
[gc]::WaitForPendingFinalizers()

# 処理完了メッセージを表示
write-host "`nReady!" -Foreground Green