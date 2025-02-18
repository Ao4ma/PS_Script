class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$IniFilePath
    [hashtable]$IniContent
    [bool]$IsLibraryConfigured
    [string]$ScriptFolder
    [string]$LogFolder
    [System.Collections.ArrayList]$ManagedInstances

    PC([string]$name, [string]$iniFilePath) {
        $this.Name = $name
        $this.IniFilePath = $iniFilePath
        $this.IPAddress = $this.GetIPAddress()
        $this.MACAddress = $this.GetMACAddress()
        $this.IniContent = $this.GetIniContent()
        $this.IsLibraryConfigured = $this.CheckLibraryConfiguration()
        $this.ManagedInstances = [System.Collections.ArrayList]::new()
    }

    [void]DisplayInfo() {
        Write-Host "PC Name: $($this.Name)"
        Write-Host "IP Address: $($this.IPAddress)"
        Write-Host "MAC Address: $($this.MACAddress)"
    }

    [string]GetIPAddress() {
        $ipConfig = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" }
        return $ipConfig.IPAddress
    }

    [string]GetMACAddress() {
        $macConfig = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        return $macConfig.MacAddress
    }

    [hashtable]GetIniContent() {
        $iniContent = @{}
        $currentSection = ""

        foreach ($line in Get-Content -Path $this.IniFilePath) {
            if ($line -match "^\[(.+)\]$") {
                $currentSection = $matches[1]
                $iniContent[$currentSection] = @{}
            } elseif ($line -match "^(.+?)=(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $iniContent[$currentSection][$key] = $value
            }
        }

        return $iniContent
    }

    [bool]CheckLibraryConfiguration() {
        # ライブラリの設定を確認するロジックをここに追加
        return $true
    }

    [void]SetScriptFolder([string]$path) {
        $this.ScriptFolder = $path
    }

    [void]SetLogFolder([string]$path) {
        $this.LogFolder = $path
    }

    [void]AddInstance([object]$instance) {
        $this.ManagedInstances.Add($instance) | Out-Null
    }

    [void]RemoveInstance([object]$instance) {
        $this.ManagedInstances.Remove($instance) | Out-Null
    }

    [void]NotifyInstanceClosed([object]$instance) {
        Write-Host "インスタンスが閉じられました: $instance"
        $this.RemoveInstance($instance)
    }

    [string]GetScriptPath([string]$libraryName) {
        $userProfile = [System.Environment]::GetFolderPath("UserProfile")
        return Join-Path -Path $userProfile -ChildPath "Documents\GitHub\PS_Script\MyLibrary\$libraryName"
    }
}