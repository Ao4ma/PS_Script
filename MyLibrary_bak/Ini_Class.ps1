class IniFile {
    [string]$FilePath
    [hashtable]$Content

    IniFile([string]$filePath) {
        $this.FilePath = $filePath
        $this.Content = $this.LoadIniFile()
    }

    [hashtable]LoadIniFile() {
        $iniContent = @{}
        $currentSection = ""

        foreach ($line in Get-Content -Path $this.FilePath) {
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

    [string]GetValue([string]$section, [string]$key) {
        if ($this.Content.ContainsKey($section) -and $this.Content[$section].ContainsKey($key)) {
            return $this.Content[$section][$key]
        } else {
            return $null
        }
    }

    [void]SetValue([string]$section, [string]$key, [string]$value) {
        if (-not $this.Content.ContainsKey($section)) {
            $this.Content[$section] = @{}
        }
        $this.Content[$section][$key] = $value
        $this.SaveIniFile()
    }

    [void]RemoveValue([string]$section, [string]$key) {
        if ($this.Content.ContainsKey($section) -and $this.Content[$section].ContainsKey($key)) {
            $this.Content[$section].Remove($key)
            if ($this.Content[$section].Count -eq 0) {
                $this.Content.Remove($section)
            }
            $this.SaveIniFile()
        }
    }

    [void]SaveIniFile() {
        $lines = @()
        foreach ($section in $this.Content.Keys) {
            $lines += "[$section]"
            foreach ($key in $this.Content[$section].Keys) {
                $lines += "$key=$($this.Content[$section][$key])"
            }
            $lines += ""
        }
        $lines | Set-Content -Path $this.FilePath
    }
}