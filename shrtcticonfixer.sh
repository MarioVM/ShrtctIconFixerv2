# Get the script's root folder
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path to your Gaming folder
$folderPath = "C:\Gaming"

# Path to the log file in the script's root folder
$logFilePath = Join-Path -Path $scriptRoot -ChildPath "icon_refresh_log.txt"

# Clear the log file if it exists
if (Test-Path $logFilePath) {
    Remove-Item $logFilePath
}

# Function to log messages to both console and log file
function Log-Message {
    param (
        [string]$message
    )
    Write-Host $message
    Add-Content -Path $logFilePath -Value $message
}

# Debugging: Check if folder exists
if (-Not (Test-Path -Path $folderPath)) {
    Log-Message "Folder path does not exist: $folderPath"
    exit
}

$shortcuts = Get-ChildItem -Path $folderPath -Filter *.lnk

foreach ($shortcut in $shortcuts) {
    $shortcutPath = $shortcut.FullName

    try {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcutObject = $wsh.CreateShortcut($shortcutPath)
        
        $targetPath = $shortcutObject.TargetPath
        $targetFolder = Split-Path -Parent $targetPath

        # Debugging: Check target path and folder
        Log-Message "Processing shortcut: $shortcutPath"
        Log-Message "Target path: $targetPath"
        Log-Message "Target folder: $targetFolder"

        # Try to find an icon file first
        $iconFile = Get-ChildItem -Path $targetFolder -Filter *.ico -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($iconFile) {
            $shortcutObject.IconLocation = $iconFile.FullName
            Log-Message "Found icon file: $($iconFile.FullName)"
        } else {
            # If no icon file, try to find an executable file
            $exeFile = Get-ChildItem -Path $targetFolder -Filter *.exe -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($exeFile) {
                $shortcutObject.IconLocation = $exeFile.FullName
                Log-Message "Found executable file: $($exeFile.FullName)"
            } else {
                $shortcutObject.IconLocation = $targetPath
                Log-Message "Using target path as icon location"
            }
        }

        $shortcutObject.Save()
        Log-Message "Refreshed icon for: $shortcutPath"
    } catch {
        Log-Message "Failed to refresh icon for: $shortcutPath"
        Log-Message $_.Exception.Message
    }
}

Log-Message "Icon refresh completed for all shortcuts."

# Wait for user input before closing the shell window
Read-Host "Press Enter to exit..."
