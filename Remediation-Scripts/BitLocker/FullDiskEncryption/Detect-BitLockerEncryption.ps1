# ------------------------------------------------------------------------------------------------------------------------------------
[string] $driveLetter = 'C:'
# ------------------------------------------------------------------------------------------------------------------------------------
enum ConversionStatus {
    FullyDecrypted = 0x0
    FullyEncrypted = 0x1
    EncryptionInProgress = 0x2
    DecryptionInProgress = 0x3
    EncryptionPaused = 0x4
    DecryptionPaused = 0x5
}

enum EncryptionFlags {
    FullEncryption = 0x0
    UsedSpaceOnly = 0x1
    WipeInProgress = 0x2
}
# ------------------------------------------------------------------------------------------------------------------------------------

# Get the encryptable (fixed) volume for defined drive letter from the associated WMI class.
$encryptableVolumes = Get-CimInstance -ClassName 'Win32_EncryptableVolume' -Namespace 'root\CIMV2\Security\MicrosoftVolumeEncryption'
$selectedDrive = $encryptableVolumes | Where-Object { $_.DriveLetter -eq $driveLetter }

# If no volume with the defined drive letter could be found, exit the script.
if (-not $selectedDrive) {
    Write-Error "Could not find an encryptable volume for drive letter $($driveLetter)"
    exit 1
}

# Try to get the conversion status of the associated drive.
$driveStatus = $selectedDrive | Invoke-CimMethod -MethodName 'GetConversionStatus'
if ($driveStatus.ExitCode -eq 1) {
    Write-Error "Could not get the conversion status for drive letter $($driveLetter)"
    exit 1
}

# If the drive doesn't report that is has fully encrypted, exit the script.
if ($driveStatus.ConversionStatus -ne [ConversionStatus]::FullyEncrypted) {
    Write-Error "Volume with drive letter $($driveLetter) has not finished encryption."
    exit 1
}

# If the drive reports an encryption flag it indicates that there wasn't full encryption selected.
# Exit the script in this case, too.
if ($driveStatus.EncryptionFlags -ne [EncryptionFlags]::FullEncryption) {
    Write-Error "Volume with drive letter $($driveLetter) is not fully encrypted."
    exit 1
}

# If nothing fails, our volume is compliant and we can exit with return code 0.
Write-Output "Volume with drive letter $($driveLetter) is compliant."
exit 0