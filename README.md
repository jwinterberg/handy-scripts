# Handy scripts

Just some handy scripts for use with Microsoft Endpoint Manager (Intune)

## Contents:
* **Remediation Scripts**
    * **BitLocker**
        * **FullDiskEncryption** (Detects whether the system disk is fully encrypted or only used space is encrypted)
* **Recreate Shortcuts**\
    > :warning: These scripts were only tested locally on my computer and not in the field. Please conduct your own tests first!
    * **Write-ShortcutsToJson**\
    Takes an OutputPath argument to specify an output JSON file location.\
    The JSON file contains the FullName (Location), TargetPath, Arguments, Description, WorkingDirectory, IconLocation and WindowStyle.
    * **New-ShortcutsFromJson**\
    Takes a Path argument to specify the JSON file from Write-ShortcutsToJson and will recreate all the shortcuts based on the data.\
    It automatically detects if a shortcut already exists or if the TargetPath exists (application not installed)
