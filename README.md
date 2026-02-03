# VMacropad

<div align="center">
  <img src="vmacropad.ico" width="128" height="128" alt="VMacropad Icon">
  <br>
  <b>An open-source, modern replacement driver for generic 3-Key / 1-Knob Macropads.</b>
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
*   Generic "3-Key Mini Keyboard" listings on Amazon/AliExpress
*   Devices that identify as "Keypad" or "USB Input Device" with the IDs above

*Note: If your device uses different IDs, you can change them directly in the app Settings menu.*

## Features
*   **Auto-Profile Switching:** Automatically changes key mappings based on which app you are using. Use Photoshop shortcuts when Photoshop is open, then switch to Media keys when Spotify is focused.
*   **App Audio Control:** Bind keys to change the volume of *specific* applications (e.g., lower Discord volume without lowering the game volume). Includes "Fuzzy Matching" so `tidal` finds `TIDALPlayer.exe`.
*   **Smart Fallback:** If the specific app you are trying to control isn't open, the knob automatically falls back to controlling the Master Volume.
*   **Modern UI:** Clean, Dark Mode interface using CustomTkinter.
*   **System Tray Integration:** Minimizes silently to the background with very low resource usage (~0% CPU).
*   **Visual Config:** Interactive visualizer to see exactly what you are programming.
*   **Instant Save:** No "Apply" button required. Changes upload to the device memory instantly.
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

# Install dependencies (Updated for v1.0.3)
pip install customtkinter hidapi pywin32 psutil pystray pillow keyboard pycaw comtypes

# Run the app
python vmacropad.py
```

## How to use Auto-Switching
1.  Create a new preset (e.g. "Photoshop").
2.  Click the **"Link to App"** button.
3.  You have 3 seconds to click on your target application window.
4.  Done. Whenever that app is in focus, the macropad will switch to that preset automatically.

## How to use App Audio Control
1.  Select a key or knob direction.
2.  Go to the **"App Audio"** tab.
3.  Click **"Grab Active App"** and click on the window you want to control (e.g., Spotify), OR manually type the process name (e.g., `spotify`).
4.  Select the action (Volume Up, Down, or Mute).
5.  Now that key will only change Spotify's volume. If Spotify is closed, it will change the system volume instead.

## Building the EXE
If you want to compile it yourself, use the following PyInstaller command. This ensures the icon, theme files, and new audio libraries are bundled correctly.

```powershell
pyinstaller --noconsole --onefile --icon="vmacropad.ico" --add-data "vmacropad.ico;." --collect-all customtkinter vmacropad.py
```

## Credits
Created by [Visiuun](https://github.com/visiuun).
Licensed under the MIT License.
