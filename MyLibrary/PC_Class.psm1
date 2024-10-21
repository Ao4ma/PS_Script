class PC {
    [string]$Name
    [string]$ScriptFolder
    [string]$LogFolder

    PC([string]$name) {
        $this.Name = $name
    }

    [void]SetScriptFolder([string]$folderPath) {
        $this.ScriptFolder = $folderPath
    }

    [void]SetLogFolder([string]$folderPath) {
        $this.LogFolder = $folderPath
    }

    [void]DisplayInfo() {
        Write-Host "PC Name: $($this.Name)"
        Write-Host "Script Folder: $($this.ScriptFolder)"
        Write-Host "Log Folder: $($this.LogFolder)"
    }

    [string] GetIPAddress() {
        $ipAddresses = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -ne '127.0.0.1' }).IPAddress
        return $ipAddresses -join ", "
    }

    [string] GetMACAddress() {
        $macAddresses = (Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }).MacAddress
        return $macAddresses -join ", "
    }

    [void] ListInstalledLibraries() {
        $libraries = Get-InstalledModule
        foreach ($library in $libraries) {
            Write-Host "$($library.Name) - $($library.Version)"
        }
    }
}