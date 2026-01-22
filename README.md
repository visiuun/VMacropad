# VMacropad

<div align="center">
  <img src="vmacropad.ico" width="128" height="128" alt="VMacropad Icon">
  <br>
  <b>An open-source, modern replacement driver for generic 3-Key / 1-Knob Macropads.</b>
  <br><br>
  <a href="https://github.com/visiuun/VMacropad/releases/latest">
    <img src="https://img.shields.io/badge/Download-Windows%20Exe-blue?style=for-the-badge&logo=windows" alt="Download">
  </a>
</div>

## Why this exists
I bought a cheap "3-key 1-knob" macropad from AliExpress/Amazon. The hardware is great, but the provided software was terrible—it looked like a Windows 98 app, flagged my antivirus, and couldn't switch profiles automatically based on the active window.

I reverse-engineered the HID protocol and built **VMacropad** to fix these issues. It's written in Python, fully transparent, and includes the features these devices should have had out of the box.

## Supported Hardware
This software is designed for macropads using the **CH57x/CH55x** chips, commonly identified by:
*   **Vendor ID:** `0x1189`
*   **Product ID:** `0x8890`

*Note: If your device has different IDs, you can override them in the `config.json` file generated after the first run.*

## Features
*   **Auto-Profile Switching:** Automatically changes key mappings based on which app you are using (e.g., Photoshop shortcuts when Photoshop is open, Media keys when Spotify is open).
*   **Modern UI:** Clean, Dark Mode interface using CustomTkinter.
*   **System Tray Integration:** Minimizes silently to the background with very low resource usage (~0% CPU).
*   **Visual Config:** Interactive visualizer to see exactly what you are programming.
*   **No "Apply" Button:** Changes upload to the device memory instantly.
*   **Portable:** Single `.exe` file, no installation required.

## Installation

### For Users (Recommended)
1.  Go to the [**Releases**](https://github.com/visiuun/VMacropad/releases) page on the right.
2.  Download `VMacropad.exe`.
3.  Run it. (You may need to "Run as Administrator" if you want it to interact with games or admin-level apps).

### For Developers (Running from Source)
If you want to modify the code or run it via Python:

```bash
# Clone the repository
git clone https://github.com/visiuun/VMacropad.git

# Install dependencies
pip install customtkinter hidapi pywin32 psutil pystray pillow

# Run the app
python vmacropad.py
```

## How to use Auto-Switching
1.  Create a preset (e.g., "Photoshop").
2.  Click **"Link to App"**.
3.  You have 3 seconds to click on your target application window.
4.  Done. Whenever that app is in focus, the macropad will switch to that preset automatically.

## Building the EXE
If you want to compile it yourself, use the following PyInstaller command to ensure the icon and theme files are bundled correctly:

```powershell
pyinstaller --noconsole --onefile --icon="vmacropad.ico" --add-data "vmacropad.ico;." --hidden-import=win32gui --hidden-import=win32process --hidden-import=psutil --name="VMacropad" --collect-all customtkinter --clean vmacropad.py
```

## Credits
Created by [Visiuun](https://github.com/visiuun).
Licensed under the MIT License.
