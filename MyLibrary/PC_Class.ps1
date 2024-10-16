class PC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$IniFilePath
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
        $this.IniContent = @{}
        if (Test-Path $this.IniFilePath) {
            $lines = Get-Content -Path $this.IniFilePath
            foreach ($line in $lines) {
                if ($line -match "^(.*)=(.*)$") {
                    $this.iniContent[$matches[1].Trim()] = $matches[2].Trim()
                }
            }
        }
        return $this.iniContent
    }

    [void]SavePCInfoToIni() {
        $this.iniContent = $this.GetIniContent()
        $this.iniContent[$this.Name] = "$($this.IPAddress),$($this.MACAddress),$($this.GetScriptDirectory())"
        $iniLines = @()
        foreach ($key in $this.iniContent.Keys) {
            $iniLines += "`"$key`",`"$($this.iniContent[$key])`""
        }
        Set-Content -Path $this.IniFilePath -Value $iniLines
    }

    [void]LoadPCInfoFromIni() {
        $this.iniContent = $this.GetIniContent()
        if ($this.iniContent.ContainsKey($this.Name)) {
            $info = $this.iniContent[$this.Name] -split ","
            $this.IPAddress = $info[0]
            $this.MACAddress = $info[1]
            $scriptDir = $info[2]
            Set-Location -Path $scriptDir
            Write-Host "Loaded PC info from INI file: $($this.Name), $($this.IPAddress), $($this.MACAddress), $scriptDir"
        } else {
            Write-Host "PC info not found in INI file for: $($this.Name)"
        }
    }

    [bool] CheckLibraryConfiguration() {
        if ($this.IniContent.ContainsKey("LibraryName") -and $this.IniContent.ContainsKey("LibraryPath")) {
            $libraryPath = $this.IniContent["LibraryPath"]
            if (Test-Path $libraryPath) {
                Add-Type -Path $libraryPath
                Write-Host "Imported Interop Assembly from $libraryPath"
                return $true
            } else {
                Write-Warning "Interop Assembly path is invalid or not found: $libraryPath"
                return $false
            }
        }
        return $false
    }
}