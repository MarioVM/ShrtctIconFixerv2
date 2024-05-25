# Path to your Gaming folder
$folderPath = "C:\Gaming"

# Refresh icon for each shortcut
$shortcuts = Get-ChildItem -Path $folderPath -Filter *.lnk

foreach ($shortcut in $shortcuts) {
    $shortcutPath = $shortcut.FullName

    try {
        # Create a WScript.Shell COM object
        $wsh = New-Object -ComObject WScript.Shell
        $shortcutObject = $wsh.CreateShortcut($shortcutPath)

        # Force icon refresh by setting and resetting the icon location
        $originalIconLocation = $shortcutObject.IconLocation
        $shortcutObject.IconLocation = "$env:windir\system32\shell32.dll,0"
        $shortcutObject.Save()
        
        # Restore the original icon location
        $shortcutObject.IconLocation = $originalIconLocation
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
