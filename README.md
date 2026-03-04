# VMacropad

<div align="center">
  <img src="vmacropad.ico" width="128" height="128" alt="VMacropad Icon">
  <br>
  <b>An open-source, modern replacement driver for generic Macropads.</b>
  <br><br>
  <a href="https://github.com/visiuun/VMacropad/releases/latest">
    <img src="https://img.shields.io/badge/Download-Windows%20Exe-blue?style=for-the-badge&logo=windows" alt="Download">
  </a>
  <br><br>
</div>

![App Screenshot](screenshot.png)

## Why this exists
I bought a cheap "3-key 1-knob" macropad from AliExpress. The hardware is great. The provided software was terrible. It looked like a Windows 98 app, flagged my antivirus, and could not switch profiles automatically based on the active window.

I reverse-engineered the HID protocol and built **VMacropad** to fix these issues. It is written in Python, fully transparent, and includes the features these devices should have had out of the box.

## Supported Hardware
This software is designed for generic macropads using the **CH57x/CH55x** chips.

**Default Hardware IDs:**
*   **Vendor ID:** `0x1189`
*   **Product ID:** `0x8890`

**Replaces software for:**
*   SayoDevice / SimPad (Generic Clones)
*   "RSoft" MacroPad
*   Generic 3-Key + Knob or 4-Key Macropad listings on Amazon/AliExpress
*   Devices that identify as "Keypad" or "USB Input Device" with the IDs above

*Note: If your device uses different IDs, you can change them directly in the app Settings menu.*

## Features
*   **Layout Support:** Now supports both **3-Key + Knob** and **4-Key** layouts. You can toggle this in the Settings menu.
*   **Auto-Profile Switching:** Automatically changes key mappings based on which app you are using.
*   **App Audio Control:** Bind keys to change the volume of *specific* applications (e.g., lower Discord volume without lowering the game volume). Includes "Fuzzy Matching" so `spotify` finds `Spotify.exe`.
*   **Automatic Updates:** Checks GitHub on startup and notifies you of new versions.
*   **Smart Fallback:** If a specifically targeted app isn't open, audio keys automatically fallback to Master Volume.
*   **Modern UI:** Clean, Dark Mode interface using CustomTkinter.
*   **System Tray Integration:** Minimizes silently to the background.
*   **Portable:** Single `.exe` file. No installation required.

## Installation

### For Users (Recommended)
1.  Go to the [**Releases**](https://github.com/visiuun/VMacropad/releases) page on the right.
2.  Download `VMacropad.exe`.
3.  Run it.

### For Developers (Running from Source)
If you want to inspect the code or run it via Python:

```bash
# Clone the repository
git clone https://github.com/visiuun/VMacropad.git

# Install dependencies (Required for Audio and Hotkey features)
pip install customtkinter hidapi pywin32 psutil pystray pillow keyboard pycaw comtypes requests

# Run the app
python vmacropad.py
```

## How to use
1.  **Layout:** Go to **Settings** and select your hardware layout ("3-Key + Knob" or "4-Key").
2.  **Auto-Switching:** Create a preset, click **"Link to App"**, and focus your target application within 3 seconds.
3.  **App Audio:** In the **"App Audio"** tab, enter the name of the process you want to control (e.g., `discord.exe`).

## Building the EXE
If you want to compile it yourself, use the following PyInstaller command. This ensures the icon, theme files, and audio libraries are bundled correctly.

```powershell
pyinstaller --noconsole --onefile --icon="vmacropad.ico" --add-data "vmacropad.ico;." --hidden-import=win32gui --hidden-import=win32process --hidden-import=psutil --hidden-import=keyboard --hidden-import=pycaw --hidden-import=comtypes --name="VMacropad" --collect-all customtkinter --clean vmacropad.py
```

## Credits
Created by [Visiuun](https://github.com/visiuun).
Licensed under the MIT License.
