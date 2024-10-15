class PC {
    [string]$Name
    [string]$IniFilePath
    [string]$IPAddress
    [string]$MACAddress
    [hashtable]$IniContent
    [bool]$IsLibraryConfigured

    PC([string]$name, [string]$iniFilePath) {
        $this.Name = $name
        $this.IniFilePath = $iniFilePath
        $this.IPAddress = $this.GetIPAddress()
        $this.MACAddress = $this.GetMACAddress()
        $this.IniContent = $this.GetIniContent()
        $this.IsLibraryConfigured = $this.CheckLibraryConfiguration()
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

    [string]GetScriptDirectory() {
        return Split-Path -Parent -Path $PSCommandPath
    }

    [void]ChangeToScriptDirectory() {
        $scriptDir = $this.GetScriptDirectory()
        Set-Location -Path $scriptDir
        Write-Host "Changed directory to script location: $scriptDir"
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
        # ライブラリの設定をチェックするロジックをここに追加
        return $true
    }
}