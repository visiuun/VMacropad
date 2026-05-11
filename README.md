# VMacropad

<div align="center">
  <img src="vmacropad.ico" width="128" height="128" alt="VMacropad Icon">
  <br>
  <b>A modern, open-source driver replacement for generic AliExpress Macropads.</b>
  <br><br>
  <a href="https://github.com/visiuun/VMacropad/releases/latest">
    <img src="https://img.shields.io/badge/Download-Latest%20Release-green?style=for-the-badge&logo=windows" alt="Download">
  </a>
  <br><br>
</div>

![App Screenshot](screenshot.png)

## Why this exists
Most "3-key 1-knob" macropads from AliExpress come with software that is essentially malware. It looks like it was made in 1998, triggers every antivirus on Earth, and doesn't support basic features like automatic profile switching or browser-aware macros.

I reverse-engineered the HID protocol to build **VMacropad**. It’s a clean, transparent Python-based manager that makes cheap hardware feel like a premium Elgato or Razer device.

## Supported Hardware
Designed for generic macropads using **CH57x/CH55x** microcontrollers.

**Default Hardware IDs:**
*   **Vendor ID:** `0x1189`
*   **Product ID:** `0x8890`

**Works with:**
*   SayoDevice / SimPad Clones
*   "RSoft" MacroPads
*   Standard 3-Key + Knob and 4-Key layouts
*   Any "USB Input Device" matching the IDs above (changeable in Settings)

## Core Features
*   **Context-Aware Auto-Switching:** Profiles change automatically based on your active window.
*   **Browser Tab Mapping:** Link presets to specific websites (YouTube, Twitch, Work) using Window Title detection. Includes **Regex** support for advanced users.
*   **Native Auto-Updater:** Updates itself directly from GitHub. No manual downloads or browser visits required.
*   **Per-App Audio Control:** Bind keys to specific processes (e.g., Knob controls Spotify volume, Keys control Discord mute) using fuzzy process matching.
*   **Modern Native UI:** Ripped out glitchy custom components for native Windows dropdowns. Supports smooth scrolling and type-to-jump navigation.
*   **Non-Admin Startup:** Runs at Windows startup without annoying UAC prompts. Admin rights are only requested when performing an actual update.

## Installation

### For Users
1.  Download `VMacropad.exe` from the [**Latest Release**](https://github.com/visiuun/VMacropad/releases).
2.  Run the portable EXE (no installation needed).

### For Developers (Source)
```bash
git clone https://github.com/visiuun/VMacropad.git
pip install customtkinter hidapi pywin32 psutil pystray pillow keyboard pycaw comtypes requests
python vmacropad.py
```

## How to use
1.  **Profiles:** Create a preset and click **"Link to App"**.
2.  **App Level:** Focus an app (e.g., Photoshop) to link the whole program.
3.  **Site Level:** Use the **"Window Title"** option to link to a browser tab keyword (e.g., `- YouTube`).
4.  **Priority:** The switcher prioritizes Window Titles (specific) > App Processes (general) > Default Preset.

## Definitive Build Command
To compile the standalone EXE with all dependencies (including the HID and Audio libraries):

```powershell
pyinstaller --noconsole --onefile --icon="vmacropad.ico" --add-data "vmacropad.ico;." --collect-all customtkinter --hidden-import psutil --hidden-import win32gui --hidden-import win32process --hidden-import win32com --hidden-import hid --hidden-import pystray --hidden-import PIL --name="VMacropad" --clean vmacropad.py
```

## Credits
Created by [Visiuun](https://github.com/visiuun). 
Built to kill the AliExpress "Software.zip" virus culture.
Licensed under MIT.
