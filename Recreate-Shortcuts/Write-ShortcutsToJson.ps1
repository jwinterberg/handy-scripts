param (
    $OutputPath = 'C:\Temp\shortcuts.json'
)


# Get me a WScript Shell object
$shellObj = New-Object -ComObject WScript.Shell

# Check the following paths for lnk files
$pathsToCheck = @('C:\Users\Public\Desktop', 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs')

$allFileNames = @()
$allShortcuts = @()

$pathsToCheck | ForEach-Object {
    # Get the full names (incl. path) of all lnk files in the paths to check recursively
    $fileNames = (Get-ChildItem -Path $_ -Filter '*.lnk' -Recurse -EA SilentlyContinue).FullName
    $allFileNames += $fileNames
}


$allFileNames | ForEach-Object {
    # Create a new shortcut object to grab the data from
    $shortCut = $shellObj.CreateShortcut($_)

    # Add to the array of all shortcuts
    $allShortcuts += @{
        'FullName' = $shortCut.FullName
        'TargetPath' = $shortCut.TargetPath
        'Arguments' = $shortCut.Arguments
        'Description' = $shortCut.Description
        'WorkingDirectory' = $shortCut.WorkingDirectory
        'IconLocation' = $shortCut.IconLocation
        'WindowStyle' = $shortCut.WindowStyle
    }
}

# Output to a JSON file
$shortCutJson = $allShortcuts | ConvertTo-Json
$shortCutJson | Out-File $OutputPath -Force
