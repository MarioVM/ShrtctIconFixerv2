# Path to your Gaming folder
$folderPath = "C:\Gaming"

$shortcuts = Get-ChildItem -Path $folderPath -Filter *.lnk

foreach ($shortcut in $shortcuts) {
    $shortcutPath = $shortcut.FullName

    try {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcutObject = $wsh.CreateShortcut($shortcutPath)
        
        $targetPath = $shortcutObject.TargetPath
        $targetFolder = Split-Path -Parent $targetPath

        # Try to find an icon file first
        $iconFile = Get-ChildItem -Path $targetFolder -Filter *.ico -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($iconFile) {
            $shortcutObject.IconLocation = $iconFile.FullName
        } else {
            # If no icon file, try to find an executable file
            $exeFile = Get-ChildItem -Path $targetFolder -Filter *.exe -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($exeFile) {
                $shortcutObject.IconLocation = $exeFile.FullName
            } else {
                $shortcutObject.IconLocation = $targetPath
            }
        }

        $shortcutObject.Save()
        Write-Host "Refreshed icon for: $shortcutPath"
    } catch {
        Write-Host "Failed to refresh icon for: $shortcutPath"
        Write-Host $_.Exception.Message
    }
}

Write-Host "Icon refresh completed for all shortcuts."

# Wait for user input before closing the shell window
Read-Host "Press Enter to exit..."
