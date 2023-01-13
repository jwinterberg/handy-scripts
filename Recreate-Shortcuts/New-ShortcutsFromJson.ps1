param (
    [string] $Path = 'C:\Temp\shortcuts.json'
)

# Get me a WScript Shell object
$shellObj = New-Object -ComObject WScript.Shell

# Get all the shortcuts from the JSON file
$shortCutJson = Get-Content -Path 'C:\temp\shortcuts.json'
$shortCuts = $shortCutJson | ConvertFrom-Json

# Create shortcuts
$shortCuts | ForEach-Object {
    Write-Output "Checking Shortcut: $($_.FullName)"
    # Check if the software is installed
    if (-not (Test-Path $_.TargetPath)) {
        Write-Output 'TargetPath does not exist!'
        return
    }

    # Check if the shortcut already exists
    if (Test-Path $_.FullName) {
        Write-Output 'Shortcut already exists!'
        return
    }

    # Get the parent directory
    $parentDir = (Split-Path $_.FullName -Parent)

    # If it doesn't already exist, create it (just in case)
    if (-not (Test-Path $parentDir)) {
        New-Item -Path $parentDir -ItemType Directory -Force
    }

    # Create new shortcut object and populate it with out data
    Write-Output 'Creating shortcut...'
    $shortcut = $shellObj.CreateShortcut($_.FullName)
    $shortcut.TargetPath = $_.TargetPath
    $shortCut.Arguments = $_.Arguments
    $shortcut.Description = $_.Description
    $shortcut.WorkingDirectory = $_.WorkingDirectory
    $shortcut.IconLocation = $_.IconLocation
    $shortcut.WindowStyle = $_.WindowStyle

    # Try to save it
    try {
        $shortcut.Save()
        Write-Output 'Success!'
    } catch {
        Write-Error "Failure! $($_.Exception.Message)"
    }
}

