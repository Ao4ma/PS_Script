class MyPC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$IniFilePath
    [hashtable]$IniContent
    [bool]$IsLibraryConfigured
    [string]$ScriptFolder
    [string]$LogFolder
    [System.Collections.ArrayList]$ManagedInstances

    MyPC([string]$name, [string]$iniFilePath) {
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
        $this.IniContent = @{}
        $currentSection = ""

        foreach ($line in Get-Content -Path $this.IniFilePath) {
            if ($line -match "^\[(.+)\]$") {
                $currentSection = $matches[1]
                $this.IniContent[$currentSection] = @{}
            } elseif ($line -match "^(.+?)=(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $this.IniContent[$currentSection][$key] = $value
            }
        }

        return $this.IniContent
    }

    [bool]CheckLibraryConfiguration() {
        Write-host "Entering CheckLibraryConfiguration method"
        Write-host "IniContent:"
        $this.IniContent.GetEnumerator() | ForEach-Object { Write-host "$($_.Key) = $($_.Value)" }
        
        try {
            Write-host "Entering try block"
            if ($this.IniContent.ContainsKey("LibraryName") -and $this.IniContent.ContainsKey("LibraryPath")) {
                Write-host "LibraryName and LibraryPath found in IniContent"
                $libraryPath = $this.IniContent["LibraryPath"]
                Write-host "LibraryPath: $libraryPath"
                if (Test-Path $libraryPath) {
                    Write-host "LibraryPath exists"
                    Add-Type -Path $libraryPath
                    Write-host "Imported Interop Assembly from $libraryPath"
                    return $true
                } else {
                    Write-Warning "Interop Assembly path is invalid or not found: $libraryPath"
                    return $false
                }
            } else {
                Write-host "LibraryName or LibraryPath not found in IniContent"
            }
        } catch {
            Write-Error "Error in CheckLibraryConfiguration: $_"
        }
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

    [void] ListInstalledLibraries() {
        $libraries = Get-InstalledModule
        foreach ($library in $libraries) {
            Write-Host "$($library.Name) - $($library.Version)"
        }
    }
}