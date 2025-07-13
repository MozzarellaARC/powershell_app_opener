function open {
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromRemainingArguments=$true)]
        [string[]]$Name
    )
    # ALIASES FOR FASTER TYPING TO OPEN APPS
    $appAliases = @{
        'vsc' = 'Visual Studio Code'
        'vscode' = 'Visual Studio Code'
        'vs'  = 'Visual Studio'
        'word' = 'Word'
        'excel' = 'Excel'
        'ppt' = 'PowerPoint'
        'ps' = 'PowerShell'
        # ADD MORE HERE
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

    if ($appMatchArray.Count -eq 1) {
        $appSelected = $appMatchArray[0]
    } else {
        Write-Host "`nAvailable matches:"
        $maxNameLen = ($appMatchArray | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
        $maxAppIdLen = ($appMatchArray | ForEach-Object { if ($_.PSObject.Properties.Name -contains 'AppID') { $_.AppID.Length } else { 9 } } | Measure-Object -Maximum).Maximum
        $i = 1
        Write-Host ("    {0,-$maxNameLen}  {1,-$maxAppIdLen}" -f 'Name', 'AppID')
        foreach ($app in $appMatchArray) {
            $appIdDisplay = if ($app.PSObject.Properties.Name -contains 'AppID') { $app.AppID } else { '<no AppID>' }
            Write-Host ("  [{0}] {1,-$maxNameLen}  {2,-$maxAppIdLen}" -f $i, $app.Name, $appIdDisplay)
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
        $appSelected = $appMatchArray | Select-Object -Skip ([int]$choice - 1) -First 1
    }

    if (-not ($appSelected.PSObject.Properties.Name -contains 'AppID')) {
        Write-Host "❌ Selected app does not have an AppID property. Object properties: $($appSelected.PSObject.Properties.Name -join ', ')" -ForegroundColor Red
        return
    }
    $appPath = $appSelected.AppID
    Write-Host ("`n🚀 Launching: {0}" -f $appSelected.Name)
    try {
        if ($appPath -match '(^[A-Z]:\\|^\\\\|[\\{.,])') {
            Start-Process "shell:AppsFolder\$appPath"
        } else {
            Start-Process "$appPath.exe"
        }
    } catch {
        Write-Host "❌ Failed to launch: $appPath" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }
}