# Cache for app list to avoid repeated Get-StartApps calls
if (-not $Script:CachedApps -or $Script:AppsCacheTime -lt (Get-Date).AddMinutes(-5)) {
    $Script:CachedApps = Get-StartApps | Sort-Object Name
    $Script:AppsCacheTime = Get-Date
}

# Function to manually refresh the app cache
function Refresh-AppCache {
    Write-Host "🔄 Refreshing app cache..." -ForegroundColor Cyan
    $Script:CachedApps = Get-StartApps | Sort-Object Name
    $Script:AppsCacheTime = Get-Date
    Write-Host "✅ App cache refreshed!" -ForegroundColor Green
}

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
    
    # Use cached apps instead of calling Get-StartApps every time
    $apps = $Script:CachedApps
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
    $userInputLower = $userInput.ToLower()
    if ($appAliases.ContainsKey($userInputLower)) {
        $searchInput = $appAliases[$userInputLower]
    }
    
    # Remove unnecessary Select-Object which copies all properties
    $appMatch = $apps | Where-Object { $_.Name -like "*$searchInput*" }
    if (-not $appMatch) {
        Write-Host "❌ No app matches input: $userInput" -ForegroundColor Red
        return
    }

    # Convert to array only if multiple matches exist
    $appMatchArray = @($appMatch)
    
    if ($appMatchArray.Count -eq 1) {
        $appSelected = $appMatchArray[0]
    } else {
        Write-Host "`nAvailable matches:"
        # Pre-calculate lengths more efficiently
        $nameWidths = $appMatchArray | ForEach-Object { $_.Name.Length }
        $maxNameLen = ($nameWidths | Measure-Object -Maximum).Maximum
        
        $appIdWidths = $appMatchArray | ForEach-Object { 
            if ($_.AppID) { $_.AppID.Length } else { 9 } 
        }
        $maxAppIdLen = ($appIdWidths | Measure-Object -Maximum).Maximum
        
        Write-Host ("    {0,-$maxNameLen}  {1,-$maxAppIdLen}" -f 'Name', 'AppID')
        for ($i = 0; $i -lt $appMatchArray.Count; $i++) {
            $app = $appMatchArray[$i]
            $appIdDisplay = if ($app.AppID) { $app.AppID } else { '<no AppID>' }
            Write-Host ("  [{0}] {1,-$maxNameLen}  {2,-$maxAppIdLen}" -f ($i + 1), $app.Name, $appIdDisplay)
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
        $appSelected = $appMatchArray[[int]$choice - 1]
    }

    # Simplified AppID check
    if (-not $appSelected.AppID) {
        Write-Host "❌ Selected app does not have an AppID property." -ForegroundColor Red
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