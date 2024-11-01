# 管理者権限で実行するためのチェック
# if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
#     Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
#     exit
# }

# 既存のWordプロセスを強制終了
$initialWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($initialWordProcesses) {
    foreach ($process in $initialWordProcesses) {
        try {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
        } catch {
            Write-Error "Failed to stop process $($process.Id): $_"
        }
    }
}

# それでもプロセスが残っている場合は、taskkillを使用
$remainingWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($remainingWordProcesses) {
    try {
        taskkill /IM WINWORD.EXE /F
    } catch {
        Write-Error "Failed to kill process using taskkill: $_"
    }
}

class WordDocument {
    [string]$DocFileName
    [string]$DocFilePath
    [string]$ScriptRoot
    [System.__ComObject]$Document
    [System.__ComObject]$WordApp

    WordDocument([string]$docFileName, [string]$docFilePath, [string]$scriptRoot) {
        $this.DocFileName = $docFileName
        $this.DocFilePath = $docFilePath
        $this.ScriptRoot = $scriptRoot
    }

    # ドキュメントを開くメソッド
    [void] OpenDocument() {
        Write-Host "IN: OpenDocument"
        try {
            $this.WordApp = New-Object -ComObject Word.Application
            $this.WordApp.Visible = $false  # 非表示モードで起動
            $this.WordApp.DisplayAlerts = 0  # ダイアログを無効にする
            $docPath = Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName
            Write-Host "Opening document at path: $docPath"

            # Wordアプリケーションが正しく起動しているか確認
            if ($null -eq $this.WordApp) {
                throw "Failed to create Word application."
            }

            # ドキュメントのパスが正しいか確認
            if (-not (Test-Path $docPath)) {
                throw "Document path does not exist: $docPath"
            }

            $this.Document = $this.WordApp.Documents.Open($docPath)
            if ($null -eq $this.Document) {
                throw "Failed to open document: $docPath"
            }
        } catch {
            Write-Error "Error in OpenDocument: $_"
            throw $_
        }
        Write-Host "OUT: OpenDocument"
    }

    # ドキュメントを閉じるメソッド
    [void] Close() {
        if ($null -ne $this.Document) {
            $this.Document.Close([ref]$false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.Document) | Out-Null
            $this.Document = $null
        }
        if ($null -ne $this.WordApp) {
            $this.WordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.WordApp) | Out-Null
            $this.WordApp = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }


    # カスタムプロパティを設定するメソッド
    [void] SetCustomProperty([string]$PropertyName, [string]$Value) {
        Write-Host "IN: SetCustomProperty"
        $customProperties = $this.Document.CustomDocumentProperties
        $binding = "System.Reflection.BindingFlags" -as [type]
        [array]$arrayArgs = $PropertyName, $false, 4, $Value
        try {
            [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            Write-Host "OUT: SetCustomProperty"
        } catch [system.exception] {
            try {
                $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $PropertyName)
                [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
                [System.__ComObject].InvokeMember("Add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
            } catch {
                Write-Error "Error while setting custom property: $_"
                throw $_
            }
            Write-Error "Error in SetCustomProperty: $_"
            throw $
        } finally {
            $this.Close()
        }
    }

    
    

    # ドキュメントを別名で保存するメソッド
    [void] SaveAs([string]$NewFilePath) {
        Write-Host "IN: SaveAs"
        try {
            $this.Document.SaveAs([ref]$NewFilePath)
            Write-Host "OUT: SaveAs"
        } catch {
            Write-Error "Error in SaveAs: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }

    # カスタムプロパティを設定して別名で保存するメソッド
    [void] SetCustomPropertyAndSaveAs([string]$PropertyName, [string]$Value) {
        Write-Host "IN: SetCustomPropertyAndSaveAs"
        try {
            $this.SetCustomProperty($PropertyName, $Value)
            $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
            $this.SaveAs($tempFilePath)
            Start-Sleep -Seconds 2  # 少し待機
            Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
            Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)
            Write-Host "OUT: SetCustomPropertyAndSaveAs"
        } catch {
            Write-Error "Error in SetCustomPropertyAndSaveAs: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }

    # プロパティを読み取るメソッド
    [string] Read_Property([string]$propName) {
        Write-Host "IN: Read_Property"
        try {
            $this.OpenDocument()
    
            $builtinProps = $this.Document.BuiltInDocumentProperties
            $customProps = $this.Document.CustomDocumentProperties
    
            $prop = $builtinProps | ForEach-Object { $_.Item($propName) } | Where-Object { $_ -ne $null }
            if ($null -eq $prop) {
                $prop = $customProps | ForEach-Object { $_.Item($propName) } | Where-Object { $_ -ne $null }
            }
    
            if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
                Write-Host "OUT: Read_Property"
                return $null
            }
    
            $propValue = $prop.Value
            Write-Host "OUT: Read_Property"
            return $propValue
        } catch {
            Write-Error "Error in Read_Property: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }
    

    # プロパティを更新するメソッド
    [void] Update_Property([string]$propName, [string]$propValue) {
        Write-Host "IN: Update_Property"
        try {
            $this.OpenDocument()

            $binding = "System.Reflection.BindingFlags" -as [type]
            $builtinProps = $this.Document.BuiltInDocumentProperties
            $customProps = $this.Document.CustomDocumentProperties

            $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $builtinProps, @($propName))
            if ($null -eq $prop) {
                $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($propName))
            }

            if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
                Write-Host "OUT: Update_Property"
                return
            }

            $this.InvokeComObjectMember($prop, "Value", "SetProperty", @($propValue))

            # 一旦別名保存し、クローズして、GCしてからファイルリネーム
            $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
            $this.SaveAs($tempFilePath)
            Start-Sleep -Seconds 2  # 少し待機
            Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
            Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)

            Write-Host "OUT: Update_Property"
        } catch {
            Write-Error "Error in Update_Property: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }

    # カスタムプロパティを削除するメソッド
    [void] Delete_Property([string]$propName) {
        Write-Host "IN: Delete_Property"
        try {
            $this.OpenDocument()
            $customProps = $this.Document.CustomDocumentProperties
            if ($this.CheckNull($customProps, "CustomDocumentProperties is null. Cannot delete property.")) {
                Write-Host "OUT: Delete_Property"
                return
            }

            $binding = "System.Reflection.BindingFlags" -as [type]
            $prop = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProps, @($propName))
            if ($this.CheckNull($prop, "Property '$($propName)' not found.")) {
                Write-Host "OUT: Delete_Property"
                return
            }

            $this.InvokeComObjectMember($customProps, "Delete", "InvokeMethod", @($propName))

            # 一旦別名保存し、クローズして、GCしてからファイルリネーム
            $tempFilePath = Join-Path -Path $this.DocFilePath -ChildPath "temp_$($this.DocFileName)"
            $this.SaveAs($tempFilePath)
            Start-Sleep -Seconds 2  # 少し待機
            Remove-Item -Path (Join-Path -Path $this.DocFilePath -ChildPath $this.DocFileName) -Force
            Rename-Item -Path $tempFilePath -NewName (Split-Path $this.DocFileName -Leaf)

            Write-Host "OUT: Delete_Property"
        } catch {
            Write-Error "Error in Delete_Property: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }

    # カスタムプロパティをファイルに書き出すメソッド
    [void] WriteCustomPropertiesToFile([string]$filePath) {
        Write-Host "IN: WriteCustomPropertiesToFile"
        try {
            $this.OpenDocument()
            $customProps = $this.Document.CustomDocumentProperties
            $output = @()

            if ($this.CheckNull($customProps, "CustomDocumentProperties is null.")) {
                $output += "Custom properties not found."
            } else {
                foreach ($prop in $customProps) {
                    $output += "$($prop.Name): $($prop.Value)"
                }
            }

            if ($output.Count -eq 0) {
                $output += "Custom properties not found."
            }

            $output | Out-File -FilePath $filePath
            Write-Host "OUT: WriteCustomPropertiesToFile"
        } catch {
            Write-Error "Error in WriteCustomPropertiesToFile: $_"
            throw $_
        } finally {
            $this.Close()
        }
    }

    # Nullチェックメソッド
    [bool] CheckNull([object]$obj, [string]$message) {
        if ($null -eq $obj) {
            Write-Host $message
            return $true
        }
        return $false
    }

    # InvokeComObjectMember メソッドを追加
    [object] InvokeComObjectMember([object]$comObject, [string]$memberName, [string]$bindingFlags, [array]$args) {
        $binding = "System.Reflection.BindingFlags" -as [type]
        return [System.__ComObject].InvokeMember($memberName, $binding::$bindingFlags, $null, $comObject, $args)
    }
}

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
Write-Host "Creating WordDocument instance"
$wordDoc = [WordDocument]::new($DocFileName, $DocFilePath, $ScriptRoot)
Write-Host "WordDocument instance created"

# メソッドの呼び出し例
$wordDoc.SetCustomPropertyAndSaveAs("CustomProperty2", "Value2")
$propValue = $wordDoc.Read_Property("CustomProperty2")
Write-Host "Read Property Value: $propValue"
$wordDoc.Update_Property("CustomProperty2", "UpdatedValue")
$wordDoc.Delete_Property("CustomProperty2")
$wordDoc.WriteCustomPropertiesToFile("C:\Users\y0927\Documents\GitHub\PS_Script\custom_properties.txt")
$wordDoc.Close()

# スクリプト終了時にWordプロセスを再度確認して強制終了
$finalWordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($finalWordProcesses) {
    foreach ($process in $finalWordProcesses) {
        try {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
        } catch {
            Write-Error "Failed to stop process $($process.Id): $_"
        }
    }
}