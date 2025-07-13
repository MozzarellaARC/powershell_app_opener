function open {
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromRemainingArguments=$true)]
        [string[]]$Name
    )
    # Alias dictionary for common abbreviations
    $appAliases = @{
        'vsc' = 'Visual Studio Code'
        'vscode' = 'Visual Studio Code'
        'vs'  = 'Visual Studio'
        'word' = 'Word'
        'excel' = 'Excel'
        'ppt' = 'PowerPoint'
        'ps' = 'PowerShell'
        # Add more as needed
    }
    $apps = Get-StartApps | Sort-Object Name
    if (-not $apps) {
        Write-Host "❌ No Start Menu apps found." -ForegroundColor Red
        return
    }
    $userInput = ($Name -join ' ')
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        Write-Host "❌ No app name provided." -ForegroundColor Yellow
        return
    }
    $searchInput = $userInput
    if ($appAliases.ContainsKey($userInput.ToLower())) {
        $searchInput = $appAliases[$userInput.ToLower()]
    }
    $appMatch = $apps | Where-Object { $_.Name -like "*$searchInput*" } | Select-Object -Property *
    $appMatchArray = @($appMatch)
    if (-not $appMatchArray -or $appMatchArray.Count -eq 0) {
        Write-Host "❌ No app matches input: $userInput" -ForegroundColor Red
        return
    }
    $appSelected = $appMatchArray[0]
    if (-not ($appSelected.PSObject.Properties.Name -contains 'AppID')) {
        Write-Host "❌ Selected app does not have an AppID property. Object properties: $($appSelected.PSObject.Properties.Name -join ', ')" -ForegroundColor Red
        return
    }
    $appPath = $appSelected.AppID
    Write-Host ("\n🚀 Launching: {0}" -f $appSelected.Name)
    try {
        if ($appPath -match '(^[A-Z]:\\|^\\\\|[\\{.,])') {
            # Full path or AppID with backslash/curly brace, use shell:AppsFolder
            Start-Process "shell:AppsFolder\$appPath"
        } else {
            # Simple AppID, not a path, not a shell id
            Start-Process "$appPath.exe"
        }
    } catch {
        Write-Host "❌ Failed to launch: $appPath" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }
}
function Open-App {
    # Alias dictionary for common abbreviations
    $appAliases = @{
        'vsc' = 'Visual Studio Code'
        'vscode' = 'Visual Studio Code'
        'vs'  = 'Visual Studio'
        'word' = 'Word'
        'excel' = 'Excel'
        'ppt' = 'PowerPoint'
        'ps' = 'PowerShell'
        # Add more as needed
    }
    $apps = Get-StartApps | Sort-Object Name
    if (-not $apps) {
        Write-Host "❌ No Start Menu apps found." -ForegroundColor Red
        return
    }

    $userInput = Read-Host "🔍 Enter part of the app name to open (or 'n' to cancel)"
    if ([string]::IsNullOrWhiteSpace($userInput) -or $userInput -match '^(n|no)$') {
        Write-Host "❌ Cancelled by user." -ForegroundColor Yellow
        return
    }
    # Use alias if available
    $searchInput = $userInput
    if ($appAliases.ContainsKey($userInput.ToLower())) {
        $searchInput = $appAliases[$userInput.ToLower()]
    }
    $appMatch = $apps | Where-Object { $_.Name -like "*$searchInput*" } | Select-Object -Property *
    $appMatchArray = @($appMatch)
    if (-not $appMatchArray -or $appMatchArray.Count -eq 0) {
        Write-Host "❌ No app matches input." -ForegroundColor Red
        return
    }
    Write-Host "\nAvailable matches:"
    $i = 1
    foreach ($app in $appMatchArray) {
        $appIdDisplay = if ($app.PSObject.Properties.Name -contains 'AppID') { $app.AppID } else { '<no AppID>' }
        Write-Host ("  [$i] {0}  (AppID: {1})" -f $app.Name, $appIdDisplay)
        $i++
    }
    $choice = Read-Host "Enter the number of the app to open, or 'n' to cancel"
    if ($choice -match '^(n|no)$') {
        Write-Host "❌ Cancelled by user." -ForegroundColor Yellow
        return
    }
    if ($choice -notmatch '^[0-9]+$' -or [int]$choice -lt 1 -or [int]$choice -gt $appMatchArray.Count) {
        Write-Host "❌ Invalid selection." -ForegroundColor Red
        return
    }
    $selected = $appMatchArray | Select-Object -Skip ([int]$choice - 1) -First 1
    if (-not ($selected.PSObject.Properties.Name -contains 'AppID')) {
        Write-Host "❌ Selected app does not have an AppID property. Object properties: $($selected.PSObject.Properties.Name -join ', ')" -ForegroundColor Red
        return
    }
    $appPath = $selected.AppID
    Write-Host ("\n🚀 Launching: {0}" -f $selected.Name)
    try {
        if ($appPath -match '(^[A-Z]:\\|^\\\\|[\\{.,])') {
            # Full path or AppID with backslash/curly brace, use shell:AppsFolder
            Start-Process "shell:AppsFolder\$appPath"
        } else {
            # Simple AppID, not a path, not a shell id
            Start-Process "$appPath.exe"
        }
    } catch {
        Write-Host "❌ Failed to launch: $appPath" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }
}
function Open-Blender {
    $blenderApps = Get-StartApps | Where-Object Name -match '^blender'
    if (-not $blenderApps) {
        Write-Host "❌ No Blender entries found in Start Menu." -ForegroundColor Red
        return
    }

    $versioned = @()
    $unversioned = @()
    foreach ($app in $blenderApps) {
        $versionStr = $app.Name -replace '^Blender\s*', ''
        try {
            $ver = [version]$versionStr
            $versioned += [PSCustomObject]@{
                Name    = $app.Name
                Version = $ver
                AppID   = $app.AppID
            }
        } catch {
            $unversioned += [PSCustomObject]@{
                Name    = $app.Name
                Version = $null
                AppID   = $app.AppID
            }
        }
    }
    $allApps = $versioned + $unversioned
    Write-Host "`n📦 Available Blender Installations:"
    $allAppsSorted = $allApps | Sort-Object { $null -ne $_.Version }, Version
    $maxNameLen = ($allAppsSorted | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
    $maxAppIdLen = ($allAppsSorted | ForEach-Object { $_.AppID.Length } | Measure-Object -Maximum).Maximum
    Write-Host ("    {0,-$maxNameLen}  {1,-$maxAppIdLen}" -f 'Name', 'AppID')
    $allAppsSorted | ForEach-Object {
        Write-Host ("    {0,-$maxNameLen}  {1,-$maxAppIdLen}" -f $_.Name, $_.AppID)
    }
    $userCin = Read-Host "`n🔍 Enter Blender version, name, or AppID (press Enter for latest version, or type 'n' to cancel)"
    if ([string]::IsNullOrWhiteSpace($userCin)) {
        if ($versioned.Count -gt 0) {
            $selected = $versioned | Sort-Object Version -Descending | Select-Object -First 1
            Write-Host "`n🚀 Launching latest version: $($selected.Name)"
        } elseif ($unversioned.Count -gt 0) {
            $selected = $unversioned | Select-Object -First 1
            Write-Host "`n🚀 Launching: $($selected.Name)"
        } else {
            Write-Host "❌ No Blender installation found." -ForegroundColor Red
            return
        }
    } elseif ($userCin -match '^(n|no)$') {
        Write-Host "❌ Cancelled by user." -ForegroundColor Yellow
        return
    } else {
        # Try version parsing
        try {
            $ver = [version]$userCin
            $selected = $versioned | Where-Object { $_.Version -eq $ver } | Select-Object -First 1
        } catch {
            $selected = $null
        }
        # Try name or AppID match if not found by version
        if (-not $selected) {
            $selected = $allAppsSorted | Where-Object { $_.Name -eq $userCin -or $_.AppID -like "*$userCin*" } | Select-Object -First 1
        }
        if (-not $selected) {
            Write-Host "❌ No Blender version, name, or AppID matches input." -ForegroundColor Red
            return
        } else {
            Write-Host "`n🚀 Launching: $($selected.Name)"
        }
    }
    $appPath = $selected.AppID
    if ($appPath -match '^[A-Z]:\\' -or $appPath -match '^\\\\') {
        Start-Process $appPath
    } else {
        Start-Process "shell:AppsFolder\$appPath"
    }
}
