# ImportModules2.ps1

# ここで直接パスを指定します
using module .\MyLibrary\WordDocumentProperties.psm1
#using module .\MyLibrary\WordDocumentUtilities.psm1
#using module .\MyLibrary\WordDocumentSignatures.psm1
using module .\MyLibrary\WordDocumentChecks.psm1
using module .\MyLibrary\WordDocument.psm1
#using module .\MyLibrary\Word_Class.psm1
using module .\MyLibrary\Word_Table.psm1
# using module .\MyLibrary\Word_Sign.psm1

# Access_Word5.psm1をインポート
# Import-Module .\MyLibrary\Access_Word_5.psm1

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

# デバッグ用設定
$docFileName = "技100-999.docx"
$scriptRoot1 = "C:\Users\y0927\Documents\GitHub\PS_Script"
$scriptRoot2 = "D:\Github\PS_Script"

# デバッグ環境に応じてパスを切り替える
if (Test-Path "D:\") {
    $scriptRoot = $scriptRoot2
} else {
    $scriptRoot = $scriptRoot1
}

$docFilePath = Join-Path -Path $scriptRoot -ChildPath $docFileName


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
    $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
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


Write-Host "Calling checkCustomProperty..."
try {
    $wordDoc.checkCustomProperty2()
    Write-Host "checkCustomProperty completed successfully."
} catch {
    Write-Error "checkCustomProperty failed: $_"
}


# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property2("CustomProperty1")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("CustomProperty2")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("CustomProperty31")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("承認者")
Write-Host "Read Property Value: $propValue"



Write-Host "Calling SetCustomPropertyAndSaveAs..."
try {
    $wordDoc.SetCustomPropertyAndSaveAs("CustomProperty31", "Value31")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("承認者", "大谷")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("承認日", "2024/11/11")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("照査者", "ベッツ")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("照査日", "2024/11/12")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("作成者", "フリーマン")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("作成日", "2024/11/13")

    Write-Host "Creating WordDocument instance..."
    # WordDocumentクラスのインスタンスを作成
    try {
        $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
        Write-Host "WordDocument instance created successfully."
    } catch {
        Write-Error "Failed to create WordDocument instance: $_"
        exit 1
    }
    $wordDoc.SetCustomPropertyAndSaveAs("CustomProperty33", "Value33")
    Write-Host "SetCustomPropertyAndSaveAs completed successfully."

} catch {
    Write-Error "SetCustomPropertyAndSaveAs failed: $_"
}

Write-Host "Creating WordDocument instance..."
# WordDocumentクラスのインスタンスを作成
try {
    $wordDoc = [WordDocument]::new($docFilePath, $scriptRoot)
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




Write-Host "Calling SetCustomProperty..."
try {
   # SetCustomProperty
   $wordDoc.SetCustomProperty("承認者", "大谷")
   $wordDoc.SetCustomProperty("承認日", "2024/11/11")
   $wordDoc.SetCustomProperty("照査者", "大谷")
   $wordDoc.SetCustomProperty("照査日", "2024/11/12")
   $wordDoc.SetCustomProperty("作成者", "フリーマン")
   $wordDoc.SetCustomProperty("作成日", "2024/11/13")
    Write-Host "SetCustomProperty completed successfully."
} catch {
    Write-Error "SetCustomProperty failed: $_"
}


# カスタムプロパティを読み取る
$propValue = $wordDoc.Read_Property2("作成者")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("作成日")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("照査者")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("照査日")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("承認者")
Write-Host "Read Property Value: $propValue"
$propValue = $wordDoc.Read_Property2("承認日")
Write-Host "Read Property Value: $propValue"












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



# 役割配列
$roles = @("承認", "照査", "作成")

# Signature_Blockクラスのインスタンスを作成
try {
    $signature_Block = [Word_Table.Signature_Block]::new($wordDoc.Document, $roles, 1)  # 直接値を使用
    Write-Host "Signature_Block インスタンスが正常に作成されました。"
} catch {
    Write-Error "エラー: $($_)"
    if ($null -ne $wordDoc) {
        $wordDoc.Close()
    }
    exit 1
}

# サイン欄の座標を取得
try {
    $signature_Coordinates = $signature_Block.Get_Signature_Coordinates()
    Write-Host "サイン欄タイプ: $($signature_Coordinates.Type)"
    foreach ($sign_Cell in $signature_Coordinates.Sign_Cells) {
        Write-Host "サイン用セル：役割: $($sign_Cell.Role), 行: $($sign_Cell.Row), 列: $($sign_Cell.Column)"
    }
} catch {
    Write-Error "エラー: $($_)"
    if ($null -ne $wordDoc) {
        $wordDoc.Close()
    }
    exit 1
}

# カスタム属性を設定
try {
    $signature_Block.Set_Custom_Attributes_at_signature_Block()
} catch {
    Write-Error "エラー: $($_)"
    if ($null -ne $wordDoc) {
        $wordDoc.Close()
    }
    exit 1
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
    Close_WordProcesses
    Write-Host "Close_Word_Processes completed successfully."
} catch {
    Write-Error "Close_Word_Processes failed: $_"
}

Write-Host "Calling Ensure_Word_Closed..."
try {
    # Wordが閉じられていることを確認する
    Ensure_WordClosed
    Write-Host "Ensure_Word_Closed completed successfully."
} catch {
    Write-Error "Ensure_Word_Closed failed: $_"
}

<#
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
#>

Write-Host "Script completed successfully."


