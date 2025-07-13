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
    $input = Read-Host "`n🔍 Enter Blender version, name, or AppID (press Enter for latest version, or type 'n' to cancel)"
    if ([string]::IsNullOrWhiteSpace($input)) {
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
    } elseif ($input -match '^(n|no)$') {
        Write-Host "❌ Cancelled by user." -ForegroundColor Yellow
        return
    } else {
        # Try version parsing
        try {
            $ver = [version]$input
            $selected = $versioned | Where-Object { $_.Version -eq $ver } | Select-Object -First 1
        } catch {
            $selected = $null
        }
        # Try name or AppID match if not found by version
        if (-not $selected) {
            $selected = $allAppsSorted | Where-Object { $_.Name -eq $input -or $_.AppID -like "*$input*" } | Select-Object -First 1
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
