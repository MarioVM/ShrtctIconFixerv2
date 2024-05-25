# Icon Refresh Script

This script refreshes the icons for all shortcut files (.lnk) in a specified folder. It attempts to set the icon to the first `.ico` file found in the target folder. If no `.ico` file is found, it will use the `.exe` file instead. If neither are found, it defaults to the target path of the shortcut.

## Prerequisites

- Python 3.x
- `pywin32` module (for interacting with Windows COM objects)

## Installation

1. Install Python 3.x from the official [Python website](https://www.python.org/).
2. Install the `pywin32` module using pip:
   ```bash
   pip install pywin32

Usage
Place the script in the desired directory.

Ensure the specified folder (e.g., C:\Gaming) contains shortcut files (.lnk).

Run the script from a command prompt or terminal:
python shrtctfixer.py

How It Works
The script gets the root directory where it is located.
It sets the path to the gaming folder (C:\Gaming).
It creates a log file in the script's root folder (icon_refresh_log.txt).
It processes each shortcut file in the gaming folder:
Gets the target path and folder of the shortcut.
Looks for .ico files in the target folder. If found, sets the icon location to the .ico file.
If no .ico file is found, looks for .exe files in the target folder. If found, sets the icon location to the .exe file.
If neither are found, sets the icon location to the target path.
Saves the changes to the shortcut and logs the actions.
The script waits for user input before closing.
Log File
All actions and any errors encountered are logged in the icon_refresh_log.txt file located in the same directory as the script.

Example
To run the script, use the following command:

python shrtctfixer.py
