import os
import win32com.client
import glob

# Get the script's root folder
script_root = os.path.dirname(os.path.abspath(__file__))

# Path to your Gaming folder
folder_path = r"C:\Gaming"

# Path to the log file in the script's root folder
log_file_path = os.path.join(script_root, "icon_refresh_log.txt")

# Clear the log file if it exists
if os.path.exists(log_file_path):
    os.remove(log_file_path)

# Function to log messages to both console and log file
def log_message(message):
    print(message)
    with open(log_file_path, "a") as log_file:
        log_file.write(message + "\n")

# Debugging: Check if folder exists
if not os.path.exists(folder_path):
    log_message(f"Folder path does not exist: {folder_path}")
    exit()

# Get all shortcuts in the folder
shortcuts = glob.glob(os.path.join(folder_path, "*.lnk"))

for shortcut_path in shortcuts:
    try:
        wsh = win32com.client.Dispatch("WScript.Shell")
        shortcut_object = wsh.CreateShortcut(shortcut_path)
        
        target_path = shortcut_object.TargetPath
        target_folder = os.path.dirname(target_path)

        # Debugging: Check target path and folder
        log_message(f"Processing shortcut: {shortcut_path}")
        log_message(f"Target path: {target_path}")
        log_message(f"Target folder: {target_folder}")

        # Try to find an icon file first
        icon_files = glob.glob(os.path.join(target_folder, "*.ico"))
        if icon_files:
            icon_file = icon_files[0]
            shortcut_object.IconLocation = icon_file
            log_message(f"Found icon file: {icon_file}")
        else:
            # If no icon file, try to find an executable file
            exe_files = glob.glob(os.path.join(target_folder, "*.exe"))
            if exe_files:
                exe_file = exe_files[0]
                shortcut_object.IconLocation = exe_file
                log_message(f"Found executable file: {exe_file}")
            else:
                shortcut_object.IconLocation = target_path
                log_message("Using target path as icon location")

        shortcut_object.Save()
        log_message(f"Refreshed icon for: {shortcut_path}")
    except Exception as e:
        log_message(f"Failed to refresh icon for: {shortcut_path}")
        log_message(str(e))

log_message("Icon refresh completed for all shortcuts.")

# Wait for user input before closing the script
input("Press Enter to exit...")
