class MyPC {
    [string]$Name
    [string]$IPAddress
    [string]$MACAddress
    [string]$IniFilePath
    [System.Collections.Generic.List[hashtable]]$IniContent
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

    [System.Collections.Generic.List[hashtable]]GetIniContent() {
        $this.iniContent = [System.Collections.Generic.List[hashtable]]::new()
        $this.currentSection = $null
        $this.currentHashTable = $null

        foreach ($line in Get-Content -Path $this.IniFilePath) {
            if ($line -match "^\[(.+)\]$") {
                if ($null -ne $this.currentSection) {
                    $this.iniContent.Add($this.currentHashTable)
                }
                $this.currentSection = $matches[1]
                $this.currentHashTable = @{ "Section" = $this.currentSection }
            } elseif ($line -match "^(.+?)=(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $this.currentHashTable[$key] = $value
            }
        }

        if ($null -ne $this.currentSection) {
            $this.iniContent.Add($this.currentHashTable)
        }

        return $this.iniContent
    }

    [bool]CheckLibraryConfiguration() {
        Write-Host "IniContent:"
        
        foreach ($section in $this.IniContent) {
            Write-Host "[$($section["Section"])]"
            foreach ($key in $section.Keys) {
                if ($key -ne "Section") {
                    Write-Host "  $key = $($section[$key])"
                }
            }
        }

        try {
            foreach ($section in $this.IniContent) {
                if ($section["Section"] -eq "LibraryPath") {
                    $libraryPath = $section["LibraryPath"]
                    if (Test-Path $libraryPath) {
                        Add-Type -Path $libraryPath
                        Write-Host "Imported Interop Assembly from $libraryPath"
                        return $true
                    } else {
                        Write-Warning "Interop Assembly path is invalid or not found: $libraryPath"
                        return $false
                    }
                }
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