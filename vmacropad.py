import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, colorchooser
import hid
import json
import os
import time
import threading
import sys
import math
import subprocess
import webbrowser
import ctypes
from PIL import Image, ImageDraw
import pystray
import re

# --- WINDOWS APP ID FIX ---
try:
    myappid = u'VMacropad.Manager.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass

# --- DEPENDENCIES CHECK ---
psutil = None
win32gui = None
win32process = None
win32com = None

try:
    import psutil
    import win32gui
    import win32process
    import win32com.client
except ImportError:
    print("Warning: 'psutil' or 'pywin32' missing. Auto-switching disabled.")

# --- THEME DEFINITION ---
class Theme:
    MAIN_BG = "#000000"
    CONTAINER_BG = "#121212" 
    WIDGET_BG = "#1c1c1c"
    ACTIVE_BUTTON = "#e5e5e5"
    INACTIVE_PILL = "#282828"
    BUTTON_HOVER = "#333333"
    CONNECTED_COLOR = "#76ff03"
    DISCONNECTED_COLOR = "#ff3d00"
    TEXT_PRIMARY = "#ffffff"
    TEXT_INVERSE = "#000000"
    TEXT_SECONDARY = "#8e8e8e"
    TEXT_DISABLED = "#4d4d4d"
    FONT_HEADER = ("Segoe UI", 20, "bold")
    FONT_SUBHEADER = ("Segoe UI", 14, "bold")
    FONT_BODY = ("Segoe UI", 12, "normal")

# --- HARDWARE DEFAULTS ---
DEFAULT_VENDOR_ID = 0x1189
DEFAULT_PRODUCT_ID = 0x8890
REPORT_ID = 0x03
ACTION_IDS = [1, 2, 3, 13, 15, 14]

# --- KEY MAPPINGS ---
KEY_MAP = {
    "None": 0,
    "A": 4, "B": 5, "C": 6, "D": 7, "E": 8, "F": 9, "G": 10,
    "H": 11, "I": 12, "J": 13, "K": 14, "L": 15, "M": 16, "N": 17, "O": 18,
    "P": 19, "Q": 20, "R": 21, "S": 22, "T": 23, "U": 24, "V": 25, "W": 26,
    "X": 27, "Y": 28, "Z": 29,
    "1": 30, "2": 31, "3": 32, "4": 33, "5": 34, "6": 35, "7": 36, "8": 37, "9": 38, "0": 39,
    "Enter": 40, "Esc": 41, "Backspace": 42, "Tab": 43, "Space": 44, "CapsLock": 57,
    "Minus (-)": 45, "Equal (=)": 46, "L Bracket ([)": 47, "R Bracket (])": 48,
    "Backslash (\\)": 49, "Non-US Hash (#)": 50, "Semicolon (;)": 51, "Quote (')": 52, 
    "Grave (`)": 53, "Comma (,)": 54, "Dot (.)": 55, "Slash (/)": 56, 
    "Intl Backslash (IS0)": 100,
    "F1": 58, "F2": 59, "F3": 60, "F4": 61, "F5": 62, "F6": 63,
    "F7": 64, "F8": 65, "F9": 66, "F10": 67, "F11": 68, "F12": 69,
    "F13": 104, "F14": 105, "F15": 106, "F16": 107, "F17": 108, "F18": 109,
    "F19": 110, "F20": 111, "F21": 112, "F22": 113, "F23": 114, "F24": 115,
    "PrintScreen": 70, "ScrollLock": 71, "Pause": 72, "Insert": 73, 
    "Home": 74, "PageUp": 75, "Delete": 76, "End": 77, "PageDown": 78,
    "Right Arrow": 79, "Left Arrow": 80, "Down Arrow": 81, "Up Arrow": 82,
    "Application / Context Menu": 101, 
    "NumLock": 83, "KP Divide (/)": 84, "KP Multiply (*)": 85, "KP Minus (-)": 86, "KP Plus (+)": 87, "KP Enter": 88,
    "KP 1": 89, "KP 2": 90, "KP 3": 91, "KP 4": 92, "KP 5": 93, "KP 6": 94,
    "KP 7": 95, "KP 8": 96, "KP 9": 97, "KP 0": 98, "KP Dot (.)": 99, "KP Equal (=)": 103,
    "Left Ctrl": 224, "Left Shift": 225, "Left Alt": 226, "Left GUI (Win)": 227,
    "Right Ctrl": 228, "Right Shift": 229, "Right Alt": 230, "Right GUI (Win)": 231,
}

MEDIA_MAP = {
    "None": (0, 0), 
    "Play/Pause": (0xCD, 0), "Stop": (0xB7, 0), 
    "Next Track": (0xB5, 0), "Prev Track": (0xB6, 0), 
    "Fast Forward": (0xB3, 0), "Rewind": (0xB4, 0),
    "Mute": (0xE2, 0), "Vol Up": (0xE9, 0), "Vol Down": (0xEA, 0),
    "Brightness Up": (0x6F, 0), "Brightness Down": (0x70, 0),
    "PC Sleep": (0x32, 0), "PC Wake": (0x33, 0),
    "Calculator": (0x92, 0x01), "File Explorer": (0x94, 0x01), 
    "Email": (0x8A, 0x01), 
    "Web Home": (0x23, 0x02), "Web Search": (0x21, 0x02),
    "Web Back": (0x24, 0x02), "Web Forward": (0x25, 0x02),
    "Web Stop": (0x26, 0x02), "Web Refresh": (0x27, 0x02),
    "Web Favorites": (0x2A, 0x02)
}

MOUSE_BUTTONS = {
    "None": 0, "Left Click": 1, "Right Click": 2, "Middle Click": 4, 
    "Back (Btn 4)": 8, "Forward (Btn 5)": 16
}

MOUSE_WHEEL = {
    "None": 0, "Scroll Up": 1, "Scroll Down": 255
}

LED_MODES = {"Off": 0, "Static": 1, "Breathing": 2}

# --- FILE PATHS ---
APP_NAME = "VMacropad"
APP_DATA_DIR = os.path.join(os.getenv('APPDATA'), APP_NAME)
if not os.path.exists(APP_DATA_DIR):
    try: os.makedirs(APP_DATA_DIR)
    except: pass
CONFIG_FILE = os.path.join(APP_DATA_DIR, "config.json")
PRESETS_FILE = os.path.join(APP_DATA_DIR, "presets.json")
MAPPINGS_FILE = os.path.join(APP_DATA_DIR, "mappings.json")

# --- HARDWARE CONTROLLER ---
class MacroPadDevice:
    def __init__(self, vendor_id, product_id):
        self.device = None
        self.working_strategy = None
        self._connected = False
        self.vid = vendor_id
        self.pid = product_id

    def is_connected(self):
        return self._connected

    def scan_for_device(self):
        try:
            devices = hid.enumerate(self.vid, self.pid)
            for d in devices:
                if d.get('interface_number') == 1:
                    return d['path']
            for d in devices:
                path_str = d['path'].decode('utf-8') if isinstance(d['path'], bytes) else d['path']
                if "mi_01" in path_str.lower():
                    return d['path']
            if devices:
                return devices[0]['path']
        except:
            pass
        return None

    def connect(self):
        if self.device:
            try: self.device.close()
            except: pass
            self.device = None
        
        target_path = self.scan_for_device()
        if target_path:
            try:
                self.device = hid.device()
                self.device.open_path(target_path)
                self.device.set_nonblocking(1)
                self._connected = True
                return True
            except:
                self.device = None
        
        self._connected = False
        return False

    def write_data(self, payload):
        if not self.device: return False
        buf_65 = [REPORT_ID] + payload + [0] * (64 - len(payload))
        strategies = [('output', buf_65), ('feature', buf_65)]
        
        if self.working_strategy == 'output': strategies = [strategies[0], strategies[1]]
        elif self.working_strategy == 'feature': strategies = [strategies[1], strategies[0]]

        for mode, buf in strategies:
            try:
                res = -1
                if mode == 'output': res = self.device.write(buf)
                else: res = self.device.send_feature_report(buf)
                if res >= 0:
                    self.working_strategy = mode
                    return True
            except: 
                continue
        return False

    def select_layer(self, layer=0): return self.write_data([0xA1, layer])
    def save_to_flash(self): return self.write_data([0xAA, 0xAA])
    
    def set_key(self, ui_index, mod, code):
        action = ACTION_IDS[ui_index]
        self.write_data([action, 1, 1, 0, mod, 0])
        return self.write_data([action, 1, 1, 1, mod, code])

    def set_media(self, ui_index, b1, b2):
        action = ACTION_IDS[ui_index]
        return self.write_data([action, 2, b1, b2])

    def set_mouse(self, ui_index, btn, scroll, mod=0):
        action = ACTION_IDS[ui_index]
        return self.write_data([action, 3, btn, 0, 0, scroll, mod])

    def set_led(self, mode): return self.write_data([0xB0, 0x08, mode])


# --- MAIN APPLICATION ---
class VMacroApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 1. Load basic configuration (Handles backward compatibility for new settings)
        self.load_config_early()
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title(f"V Macropad Manager")
        self.geometry("1000x700")
        self.configure(fg_color=Theme.MAIN_BG)
        self.protocol("WM_DELETE_WINDOW", self.on_close_attempt)
        self.minsize(950, 650)
        
        try:
            icon_path = self.resource_path("vmacropad.ico")
            self.iconbitmap(icon_path)
        except Exception:
            pass

        self.pad = MacroPadDevice(self.cfg_vid, self.cfg_pid)
        self.presets = self.load_presets()
        self.app_mappings = self.load_mappings()
        
        self.current_data = [{"type": "key", "mod": 0, "code": 0, "mouse_btn": 0, "mouse_scroll": 0} for _ in range(6)]
        self.led_mode = 1
        self.selected_key_index = 0
        self.current_preset_name = None
        
        self.is_uploading = False
        self.upload_lock = threading.Lock()
        self.connected_last_frame = False
        self.tray_icon = None
        self.running = True

        self.last_detected_target = None
        self.focus_timer_start = 0
        self.last_auto_uploaded_preset = None
        
        # New flag: If user manually selects a preset, pause auto-switching until a LINKED app is found
        self.manual_override = False

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_sidebar()
        self.setup_main_area()
        
        # 2. Load the UI state vars (current preset name from config)
        self.load_config_state_ui_vars()
        
        # 3. NOW it is safe to save. This updates old config files to new format automatically.
        self.save_config_state()
        
        self.refresh_preset_list() 
        self.toggle_startup()

        self.perform_initial_connection()
        self.setup_tray()
        
        self.check_conn_loop()
        self.app_monitor_loop()
        
        self.init_complete = True
        
        if self.cfg_tray_enabled:
            self.withdraw() 
        else:
            self.deiconify()

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def load_config_early(self):
        # Set defaults first
        self.cfg_vid = DEFAULT_VENDOR_ID
        self.cfg_pid = DEFAULT_PRODUCT_ID
        self.default_preset_name = None
        self.cfg_notify_preset = False
        self.cfg_notify_status = False
        self.cfg_tray_enabled = True
        self.cfg_startup = True
        self.cfg_focus_delay = 0.5  # Default Fast Response
        
        # Attempt to load from file
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    conf = json.load(f)
                    self.cfg_vid = conf.get("vendor_id", DEFAULT_VENDOR_ID)
                    self.cfg_pid = conf.get("product_id", DEFAULT_PRODUCT_ID)
                    self.default_preset_name = conf.get("default_preset")
                    self.cfg_notify_preset = conf.get("notify_preset", False)
                    self.cfg_notify_status = conf.get("notify_status", False)
                    self.cfg_tray_enabled = conf.get("tray_enabled", True)
                    self.cfg_startup = conf.get("startup_enabled", True)
                    # Backward compatibility: .get() returns 0.5 if key is missing
                    self.cfg_focus_delay = conf.get("focus_delay", 0.5) 
            except: pass

    def load_config_state_ui_vars(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    conf = json.load(f)
                    last = conf.get("last_preset")
                    if last and last in self.presets:
                        self.load_preset_by_name(last)
            except: pass
        
        if not self.current_preset_name and self.presets:
            self.load_preset_by_name(list(self.presets.keys())[0])

    def save_config_state(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump({
                    "vendor_id": self.cfg_vid,
                    "product_id": self.cfg_pid,
                    "last_preset": self.current_preset_name,
                    "default_preset": self.default_preset_name,
                    "notify_preset": self.cfg_notify_preset,
                    "notify_status": self.cfg_notify_status,
                    "tray_enabled": self.cfg_tray_enabled,
                    "startup_enabled": self.cfg_startup,
                    "focus_delay": self.cfg_focus_delay
                }, f, indent=4)
        except: pass

    def perform_initial_connection(self):
        def attempt():
            for _ in range(10):
                if not self.running: return
                if self.pad.connect():
                    self.after(0, self.safe_update_status, True)
                    return
                time.sleep(0.5)
        threading.Thread(target=attempt, daemon=True).start()

    def safe_update_status(self, connected):
        if self.running and self.winfo_exists():
            self.update_status_ui(connected)

    def notify_user(self, title, message):
        if self.tray_icon:
            self.update_tray_icon()
            def _delayed_notify():
                time.sleep(0.25)
                if self.tray_icon:
                    self.tray_icon.notify(message, title)
            threading.Thread(target=_delayed_notify, daemon=True).start()

    def load_presets(self):
        if os.path.exists(PRESETS_FILE):
            try:
                with open(PRESETS_FILE, "r") as f: return json.load(f)
            except: pass
        return {}

    def save_presets_file(self):
        try:
            with open(PRESETS_FILE, "w") as f: json.dump(self.presets, f, indent=4)
        except: pass

    def load_mappings(self):
        if os.path.exists(MAPPINGS_FILE):
            try:
                with open(MAPPINGS_FILE, "r") as f: return json.load(f)
            except: pass
        return {}

    def save_mappings_file(self):
        try:
            with open(MAPPINGS_FILE, "w") as f: json.dump(self.app_mappings, f, indent=4)
        except: pass

    def toggle_startup(self):
        try:
            startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
            shortcut_path = os.path.join(startup_folder, "VMacropad.lnk")
            
            if self.cfg_startup:
                target = sys.executable
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortcut(shortcut_path)
                shortcut.TargetPath = target
                if not getattr(sys, 'frozen', False):
                    shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
                shortcut.WorkingDirectory = os.path.dirname(target)
                shortcut.IconLocation = target
                shortcut.Save()
            else:
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
        except Exception: pass

    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, fg_color=Theme.CONTAINER_BG, corner_radius=0, width=260)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(2, weight=1)
        self.sidebar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.sidebar, text="V MACROPAD", font=Theme.FONT_HEADER, text_color=Theme.TEXT_PRIMARY).grid(row=0, column=0, pady=(30, 10))
        ctk.CTkLabel(self.sidebar, text="PRESET LIBRARY", font=Theme.FONT_BODY, text_color=Theme.TEXT_SECONDARY).grid(row=1, column=0, pady=(0, 20))

        self.preset_scroll = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.preset_scroll.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        self.preset_widgets = {} 

        btn_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", pady=20, padx=20)
        btn_frame.grid_columnconfigure((0,1), weight=1)

        self.btn_add = ctk.CTkButton(btn_frame, text="NEW", font=Theme.FONT_BODY, fg_color=Theme.BUTTON_HOVER, hover_color=Theme.TEXT_DISABLED, command=self.add_preset)
        self.btn_add.grid(row=0, column=0, padx=5, sticky="ew")
        
        self.btn_del = ctk.CTkButton(btn_frame, text="DELETE", font=Theme.FONT_BODY, fg_color="#441111", hover_color="#802122", command=self.del_preset)
        self.btn_del.grid(row=0, column=1, padx=5, sticky="ew")

        self.btn_settings = ctk.CTkButton(self.sidebar, text="SETTINGS", font=Theme.FONT_BODY, fg_color="transparent", border_width=1, border_color=Theme.TEXT_DISABLED, command=self.open_settings_ui)
        self.btn_settings.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 20))
        
        self.refresh_preset_list()

    def open_settings_ui(self):
        win = ctk.CTkToplevel(self)
        win.title("Settings")
        win.geometry("420x650")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.configure(fg_color=Theme.CONTAINER_BG)
        
        try:
            icon_path = self.resource_path("vmacropad.ico")
            win.iconbitmap(icon_path)
        except: pass
        
        try:
            x = self.winfo_x() + (self.winfo_width()//2) - 210
            y = self.winfo_y() + (self.winfo_height()//2) - 325
            win.geometry(f"+{x}+{y}")
        except: pass

        ctk.CTkLabel(win, text="SETTINGS", font=Theme.FONT_HEADER).pack(pady=(20, 20))
        
        var_notif_p = ctk.BooleanVar(value=self.cfg_notify_preset)
        var_notif_s = ctk.BooleanVar(value=self.cfg_notify_status)
        var_tray = ctk.BooleanVar(value=self.cfg_tray_enabled)
        var_start = ctk.BooleanVar(value=self.cfg_startup)
        
        frm_hw = ctk.CTkFrame(win, fg_color="transparent")
        frm_hw.pack(pady=10, padx=40, fill="x")
        ctk.CTkLabel(frm_hw, text="Vendor ID (Hex):", text_color=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w")
        entry_vid = ctk.CTkEntry(frm_hw, width=120)
        entry_vid.insert(0, hex(self.cfg_vid))
        entry_vid.grid(row=0, column=1, padx=10)
        
        ctk.CTkLabel(frm_hw, text="Product ID (Hex):", text_color=Theme.TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=5)
        entry_pid = ctk.CTkEntry(frm_hw, width=120)
        entry_pid.insert(0, hex(self.cfg_pid))
        entry_pid.grid(row=1, column=1, padx=10, pady=5)

        # --- Lifespan Strategy UI ---
        frm_life = ctk.CTkFrame(win, fg_color="transparent")
        frm_life.pack(pady=15, padx=40, fill="x")
        ctk.CTkLabel(frm_life, text="Auto-Switch Strategy:", text_color=Theme.TEXT_SECONDARY, font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        
        # Initial value logic
        init_strategy = "Fast Response (0.5s)"
        if self.cfg_focus_delay > 1.0:
            init_strategy = "Max Lifespan (2.0s)"
            
        seg_strategy = ctk.CTkSegmentedButton(frm_life, values=["Fast Response (0.5s)", "Max Lifespan (2.0s)"])
        seg_strategy.set(init_strategy)
        seg_strategy.pack(fill="x")

        def save_and_close():
            self.cfg_notify_preset = var_notif_p.get()
            self.cfg_notify_status = var_notif_s.get()
            old_tray = self.cfg_tray_enabled
            self.cfg_tray_enabled = var_tray.get()
            if old_tray != self.cfg_tray_enabled:
                if self.cfg_tray_enabled: self.setup_tray()
                else: 
                    if self.tray_icon: self.tray_icon.stop()
                    self.tray_icon = None
            self.cfg_startup = var_start.get()
            
            # Save strategy logic
            strat = seg_strategy.get()
            if "Fast" in strat:
                self.cfg_focus_delay = 0.5
            else:
                self.cfg_focus_delay = 2.0
            
            self.toggle_startup()
            try:
                new_vid = int(entry_vid.get(), 16)
                new_pid = int(entry_pid.get(), 16)
                if new_vid != self.cfg_vid or new_pid != self.cfg_pid:
                    self.cfg_vid = new_vid
                    self.cfg_pid = new_pid
                    self.pad.vid = new_vid
                    self.pad.pid = new_pid
                    messagebox.showinfo("Hardware IDs", "Hardware IDs updated. Reconnecting...")
                    self.pad.connect()
            except ValueError:
                messagebox.showerror("Error", "Invalid Hex format for VID/PID (Use 0x...)")
                return
            self.save_config_state()
            win.destroy()

        ctk.CTkCheckBox(win, text="Notify on Preset Change", variable=var_notif_p).pack(pady=10, padx=40, anchor="w")
        ctk.CTkCheckBox(win, text="Notify on Connect/Disconnect", variable=var_notif_s).pack(pady=10, padx=40, anchor="w")
        ctk.CTkCheckBox(win, text="Enable System Tray Icon", variable=var_tray).pack(pady=10, padx=40, anchor="w")
        ctk.CTkCheckBox(win, text="Start with Windows", variable=var_start).pack(pady=10, padx=40, anchor="w")
        
        ctk.CTkButton(win, text="SAVE & CLOSE", command=save_and_close, fg_color=Theme.ACTIVE_BUTTON, text_color="black").pack(pady=30)
        
        ctk.CTkLabel(win, text="Created by Visiuun", text_color=Theme.TEXT_SECONDARY).pack(pady=(10, 0))
        link = ctk.CTkButton(win, text="github.com/visiuun", fg_color="transparent", text_color="#4da6ff", hover=False, 
                             command=lambda: webbrowser.open("https://github.com/visiuun"))
        link.pack()

    def setup_main_area(self):
        self.main_frame = ctk.CTkFrame(self, fg_color=Theme.MAIN_BG, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_rowconfigure(1, weight=1) 
        self.main_frame.grid_columnconfigure(0, weight=1)

        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(10, 20))
        
        self.lbl_status_icon = ctk.CTkLabel(header_frame, text="●", font=("Arial", 24), text_color=Theme.DISCONNECTED_COLOR)
        self.lbl_status_icon.pack(side="left", padx=(0, 10))
        
        self.lbl_status_text = ctk.CTkLabel(header_frame, text="DISCONNECTED", font=Theme.FONT_SUBHEADER, text_color=Theme.TEXT_SECONDARY)
        self.lbl_status_text.pack(side="left")

        self.btn_set_default = ctk.CTkButton(header_frame, text="SET AS DEFAULT", 
                                            font=("Segoe UI", 11, "bold"), 
                                            fg_color="#333", 
                                            width=120, 
                                            command=self.set_active_as_default)
        self.btn_set_default.pack(side="right", padx=5)

        self.btn_link_app = ctk.CTkButton(header_frame, text="LINK TO APP", 
                                        font=("Segoe UI", 11, "bold"), 
                                        fg_color=Theme.WIDGET_BG, 
                                        width=120, 
                                        command=self.link_current_preset_to_app)
        self.btn_link_app.pack(side="right", padx=5)

        self.vis_container = ctk.CTkFrame(self.main_frame, fg_color=Theme.WIDGET_BG, corner_radius=15)
        self.vis_container.grid(row=1, column=0, sticky="nsew", pady=10)
        
        self.canvas = tk.Canvas(self.vis_container, bg=Theme.WIDGET_BG, highlightthickness=0, height=250)
        self.canvas.pack(fill="both", expand=True, padx=20, pady=20)
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<Configure>", lambda e: self.draw_visualizer())

        self.editor_frame = ctk.CTkTabview(self.main_frame, fg_color=Theme.CONTAINER_BG, 
                                         text_color=Theme.TEXT_PRIMARY,
                                         segmented_button_fg_color=Theme.WIDGET_BG,
                                         segmented_button_selected_color=Theme.ACTIVE_BUTTON,
                                         segmented_button_selected_hover_color=Theme.ACTIVE_BUTTON,
                                         segmented_button_unselected_color=Theme.WIDGET_BG,
                                         segmented_button_unselected_hover_color=Theme.BUTTON_HOVER)
        
        self.editor_frame.grid(row=2, column=0, sticky="ew", pady=20)
        
        self.tab_input = self.editor_frame.add("Input / Macro") 
        self.tab_media = self.editor_frame.add("Media")
        self.tab_led = self.editor_frame.add("LED")
        self.tab_mappings = self.editor_frame.add("App Mappings")
        
        self.setup_tab_content()

        self.btn_upload = ctk.CTkButton(self.main_frame, text="UPLOAD CONFIGURATION", 
                                      font=Theme.FONT_SUBHEADER,
                                      height=50,
                                      corner_radius=8,
                                      fg_color=Theme.WIDGET_BG,
                                      text_color=Theme.TEXT_DISABLED,
                                      state="disabled",
                                      command=self.start_upload)
        self.btn_upload.grid(row=3, column=0, sticky="ew", pady=(10,0))

    def setup_tab_content(self):
        input_container = ctk.CTkFrame(self.tab_input, fg_color="transparent")
        input_container.pack(pady=5, fill="both", expand=True)

        ctk.CTkLabel(input_container, text="Modifiers (Apply to Key & Mouse)", font=("Segoe UI", 12, "bold"), text_color=Theme.TEXT_SECONDARY).pack(pady=(5,0))
        mod_frame = ctk.CTkFrame(input_container, fg_color="transparent")
        mod_frame.pack(pady=5)
        self.var_ctrl = ctk.BooleanVar()
        self.var_shift = ctk.BooleanVar()
        self.var_alt = ctk.BooleanVar()
        self.var_win = ctk.BooleanVar()
        
        for t, v in [("Ctrl", self.var_ctrl), ("Shift", self.var_shift), ("Alt", self.var_alt), ("Win", self.var_win)]:
            ctk.CTkCheckBox(mod_frame, text=t, variable=v, command=self.store_ui_state, fg_color=Theme.ACTIVE_BUTTON, text_color=Theme.TEXT_PRIMARY).pack(side="left", padx=10)

        ctk.CTkLabel(input_container, text="Keyboard Key", font=("Segoe UI", 12, "bold"), text_color=Theme.TEXT_SECONDARY).pack(pady=(15,0))
        self.cb_key = ctk.CTkComboBox(input_container, values=list(KEY_MAP.keys()), command=self.store_ui_state, width=300)
        self.cb_key.pack(pady=5)

        ctk.CTkLabel(input_container, text="--- AND / OR ---", font=("Segoe UI", 10), text_color=Theme.TEXT_DISABLED).pack(pady=5)

        mouse_frame = ctk.CTkFrame(input_container, fg_color="transparent")
        mouse_frame.pack(pady=5)
        
        ctk.CTkLabel(mouse_frame, text="Mouse Button", font=("Segoe UI", 12, "bold"), text_color=Theme.TEXT_SECONDARY).grid(row=0, column=0, padx=10)
        self.cb_mouse_btn = ctk.CTkComboBox(mouse_frame, values=list(MOUSE_BUTTONS.keys()), command=self.store_ui_state, width=140)
        self.cb_mouse_btn.grid(row=1, column=0, padx=10)

        ctk.CTkLabel(mouse_frame, text="Mouse Wheel", font=("Segoe UI", 12, "bold"), text_color=Theme.TEXT_SECONDARY).grid(row=0, column=1, padx=10)
        self.cb_mouse_scroll = ctk.CTkComboBox(mouse_frame, values=list(MOUSE_WHEEL.keys()), command=self.store_ui_state, width=140)
        self.cb_mouse_scroll.grid(row=1, column=1, padx=10)

        self.cb_media = ctk.CTkComboBox(self.tab_media, values=list(MEDIA_MAP.keys()), command=self.store_ui_state, width=300)
        self.cb_media.pack(pady=30)

        self.cb_led = ctk.CTkComboBox(self.tab_led, values=list(LED_MODES.keys()), command=self.store_led_state, width=300)
        self.cb_led.pack(pady=30)

        self.mapping_scroll = ctk.CTkScrollableFrame(self.tab_mappings, fg_color="transparent")
        self.mapping_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        self.refresh_mappings_ui()

    def set_active_as_default(self):
        if not self.current_preset_name: return
        self.default_preset_name = self.current_preset_name
        self.app_mappings = {k: v for k, v in self.app_mappings.items() if v != self.default_preset_name}
        self.save_mappings_file()
        self.save_config_state()
        self.refresh_preset_list()
        self.refresh_mappings_ui()
        messagebox.showinfo("Success", f"'{self.default_preset_name}' is now the Default.")

    def refresh_mappings_ui(self):
        if not self.running or not self.winfo_exists(): return
        for w in self.mapping_scroll.winfo_children(): w.destroy()
        if not self.app_mappings:
            ctk.CTkLabel(self.mapping_scroll, text="No app mappings yet.", text_color=Theme.TEXT_SECONDARY).pack(pady=20)
            return
        for app_name, preset_name in list(self.app_mappings.items()):
            row = ctk.CTkFrame(self.mapping_scroll, fg_color=Theme.WIDGET_BG)
            row.pack(fill="x", pady=2, padx=5)
            ctk.CTkLabel(row, text=app_name, width=200, anchor="w", font=("Segoe UI", 12, "bold")).pack(side="left", padx=10)
            ctk.CTkLabel(row, text=f"→  {preset_name}", text_color=Theme.TEXT_SECONDARY).pack(side="left", padx=10)
            ctk.CTkButton(row, text="X", width=30, fg_color="#441111", command=lambda a=app_name: self.remove_mapping(a)).pack(side="right", padx=5)

    def remove_mapping(self, app_name):
        if app_name in self.app_mappings:
            del self.app_mappings[app_name]
            self.save_mappings_file()
            self.refresh_mappings_ui()

    def link_current_preset_to_app(self):
        if not self.current_preset_name: return
        if self.current_preset_name == self.default_preset_name:
            messagebox.showerror("Error", "The Default Preset cannot be linked to a specific app.")
            return
        def delayed_capture():
            time.sleep(3)
            app_name = self.get_active_app_process(force=True)
            if app_name and app_name.lower() not in ["python.exe", "vmacropad.exe"]:
                self.app_mappings[app_name] = self.current_preset_name
                self.save_mappings_file()
                self.after(0, self.refresh_mappings_ui)
                self.after(0, lambda: messagebox.showinfo("Success", f"Linked '{app_name}' to '{self.current_preset_name}'"))
            else:
                self.after(0, lambda: messagebox.showerror("Error", "Could not identify app."))
        messagebox.showinfo("Ready", "Focus target app within 3 seconds.")
        threading.Thread(target=delayed_capture, daemon=True).start()

    def get_active_app_process(self, force=False):
        if not psutil: return None
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return psutil.Process(pid).name()
        except:
            return None

    def app_monitor_loop(self):
        if not self.running: return
        
        current_app = self.get_active_app_process()
        
        # Determine the intended preset for the current app
        target_preset = None
        
        # Case 1: App is linked
        if current_app and current_app in self.app_mappings:
            target_preset = self.app_mappings[current_app]
            # If we hit a linked app, we disable any manual override
            if self.manual_override:
                self.manual_override = False
        
        # Case 2: App is not linked (or no app detected)
        else:
            # If manual override is active, stick to current, don't change anything
            if self.manual_override:
                target_preset = self.last_auto_uploaded_preset # Maintain status quo
            else:
                # Standard behavior: Fallback to default
                target_preset = self.default_preset_name

        # Apply logic
        if target_preset:
            if target_preset != self.last_detected_target:
                self.last_detected_target = target_preset
                self.focus_timer_start = time.time()
            else:
                if (time.time() - self.focus_timer_start) > self.cfg_focus_delay:
                    if target_preset != self.last_auto_uploaded_preset:
                        if self.pad.is_connected() and not self.is_uploading:
                            self.last_auto_uploaded_preset = target_preset
                            self.after(0, self.safe_auto_load, target_preset)
        
        self.after(1000, self.app_monitor_loop)

    def safe_auto_load(self, target_preset):
        if self.running and self.winfo_exists():
            # is_auto=True prevents setting manual_override to True
            self.load_preset_by_name(target_preset, is_auto=True)
            self.start_upload()

    def load_preset_by_name(self, name, is_auto=False):
        if name not in self.presets: return
        
        # If loaded manually, enable override lock so we don't snap back to Default immediately
        if not is_auto:
            self.manual_override = True
            # Also update last_auto_uploaded so the loop knows where we are
            self.last_auto_uploaded_preset = name
            
        self.current_preset_name = name
        self.save_config_state()
        data = self.presets[name]
        
        cleaned_data = []
        for d in data["keys"]:
            new_d = dict(d)
            if "type" not in new_d: new_d["type"] = "key"
            if "code" not in new_d: new_d["code"] = 0
            if "mouse_btn" not in new_d: new_d["mouse_btn"] = 0
            if "mouse_scroll" not in new_d: new_d["mouse_scroll"] = 0
            
            if new_d["type"] == "mouse":
                if "btn" in new_d: new_d["mouse_btn"] = new_d["btn"]
                if "scroll" in new_d: new_d["mouse_scroll"] = new_d["scroll"]

            cleaned_data.append(new_d)
            
        self.current_data = cleaned_data
        self.led_mode = data.get("led", 1)
        
        if self.winfo_exists():
            self.refresh_preset_list_highlight()
            self.update_editor_ui()
            self.draw_visualizer()
            self.update_tray_icon() 
        
        if self.init_complete and self.cfg_notify_preset:
            msg = f"{name}"
            if is_auto: msg += " (Auto)"
            self.notify_user("Preset Changed", msg)

    def refresh_preset_list(self):
        if not self.running or not self.winfo_exists(): return
        for w in self.preset_scroll.winfo_children(): w.destroy()
        self.preset_widgets.clear()
        for name in self.presets:
            display_text = f"{name} (DEFAULT)" if name == self.default_preset_name else name
            btn = ctk.CTkButton(self.preset_scroll, text=display_text, font=Theme.FONT_BODY, height=35, fg_color=Theme.INACTIVE_PILL, hover_color=Theme.BUTTON_HOVER, command=lambda n=name: self.load_preset_by_name(n))
            btn.pack(fill="x", pady=2)
            self.preset_widgets[name] = btn
        self.refresh_preset_list_highlight()

    def refresh_preset_list_highlight(self):
        if not self.running or not self.winfo_exists(): return
        for name, btn in self.preset_widgets.items():
            try:
                if name == self.current_preset_name:
                    btn.configure(fg_color=Theme.ACTIVE_BUTTON, text_color=Theme.TEXT_INVERSE)
                else:
                    btn.configure(fg_color=Theme.INACTIVE_PILL, text_color=Theme.TEXT_PRIMARY)
            except: pass

    def create_rounded_rect(self, x1, y1, x2, y2, radius=25, **kwargs):
        points = [x1+radius, y1, x1+radius, y1, x2-radius, y1, x2-radius, y1, x2, y1, x2, y1+radius, x2, y1+radius, x2, y2-radius, x2, y2-radius, x2, y2, x2-radius, y2, x2-radius, y2, x1+radius, y2, x1+radius, y2, x1, y2, x1, y2-radius, x1, y2-radius, x1, y1+radius, x1, y1+radius, x1, y1]
        return self.canvas.create_polygon(points, **kwargs, smooth=True)

    def draw_visualizer(self):
        if not self.running or not self.winfo_exists(): return
        self.canvas.delete("all")
        accent = "#888888"
        if self.current_preset_name in self.presets:
            accent = self.presets[self.current_preset_name].get("color", "#888888")
        
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 10: return
        
        cx, cy = w // 2, h // 2
        key_size = 80
        gap = 30
        
        total_width = (3 * key_size) + (3 * gap) + 120
        start_x = cx - (total_width / 2)
        key_y = cy - (key_size // 2)
        
        for i in range(3):
            x = start_x + (i * (key_size + gap))
            is_sel = (i == self.selected_key_index)
            fill = accent if is_sel else Theme.CONTAINER_BG
            outline = "#ffffff" if is_sel else "#333333"
            width = 3 if is_sel else 2
            
            tag = f"key_{i}"
            self.create_rounded_rect(x, key_y, x+key_size, key_y+key_size, radius=15, fill=fill, outline=outline, width=width, tags=tag)
            text_color = Theme.TEXT_INVERSE if (is_sel and not self.is_dark(accent)) else Theme.TEXT_PRIMARY
            self.canvas.create_text(x + key_size/2, key_y + key_size/2, text=str(i+1), fill=text_color, font=("Segoe UI", 24, "bold"), tags=tag)

        knob_x = start_x + (3 * (key_size + gap)) + 60
        knob_y = cy
        knob_r = 50
        is_ccw = (self.selected_key_index == 3)
        is_cw = (self.selected_key_index == 4)
        is_press = (self.selected_key_index == 5)
        
        self.canvas.create_oval(knob_x-knob_r, knob_y-knob_r, knob_x+knob_r, knob_y+knob_r, fill=Theme.CONTAINER_BG, outline="#333333", width=2)
        self.canvas.create_arc(knob_x-knob_r, knob_y-knob_r, knob_x+knob_r, knob_y+knob_r, start=90, extent=180, fill=accent if is_ccw else "#444444", style=tk.PIESLICE)
        self.canvas.create_arc(knob_x-knob_r, knob_y-knob_r, knob_x+knob_r, knob_y+knob_r, start=270, extent=180, fill=accent if is_cw else "#444444", style=tk.PIESLICE)
        self.canvas.create_oval(knob_x-25, knob_y-25, knob_x+25, knob_y+25, fill=Theme.CONTAINER_BG, outline="#222")
        self.canvas.create_oval(knob_x-18, knob_y-18, knob_x+18, knob_y+18, fill=accent if is_press else "#222222", outline="white" if is_press else "#555")
        self.canvas.create_text(knob_x-65, knob_y, text="CCW", fill=Theme.TEXT_SECONDARY, font=("Segoe UI", 10, "bold"), anchor="e")
        self.canvas.create_text(knob_x+65, knob_y, text="CW", fill=Theme.TEXT_SECONDARY, font=("Segoe UI", 10, "bold"), anchor="w")

    def is_dark(self, hex_color):
        if not hex_color.startswith('#'): return True
        h = hex_color.lstrip('#')
        try:
            rgb = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
            return (rgb[0]*0.299 + rgb[1]*0.587 + rgb[2]*0.114) < 140
        except: return True

    def on_canvas_click(self, event):
        if self.is_uploading or not self.winfo_exists(): return
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        cx, cy = w // 2, h // 2
        key_size = 80
        gap = 30
        
        total_width = (3 * key_size) + (3 * gap) + 120
        start_x = cx - (total_width / 2)
        key_y = cy - (key_size // 2)
        
        for i in range(3):
            kx = start_x + (i * (key_size + gap))
            if kx <= event.x <= kx+key_size and key_y <= event.y <= key_y+key_size:
                self.selected_key_index = i
                self.update_editor_ui()
                self.draw_visualizer()
                return
        
        knob_x = start_x + (3 * (key_size + gap)) + 60
        knob_y = cy
        dx = event.x - knob_x
        dy = event.y - knob_y
        dist = math.sqrt(dx*dx + dy*dy)
        if dist <= 18:
            self.selected_key_index = 5
        elif dist <= 50:
            self.selected_key_index = 3 if dx < 0 else 4
        
        self.update_editor_ui()
        self.draw_visualizer()

    def update_editor_ui(self):
        if not self.running or not self.winfo_exists(): return
        d = self.current_data[self.selected_key_index]
        dtype = d.get("type", "key")
        
        try:
            if dtype == "media":
                self.editor_frame.set("Media")
                self.cb_media.set(next((k for k,v in MEDIA_MAP.items() if v == (d.get("b1", 0), d.get("b2", 0))), "None"))
            else:
                self.editor_frame.set("Input / Macro")
                mod = d.get("mod", 0)
                self.var_ctrl.set(bool(mod & 1))
                self.var_shift.set(bool(mod & 2))
                self.var_alt.set(bool(mod & 4))
                self.var_win.set(bool(mod & 8))
                
                code = d.get("code", 0)
                self.cb_key.set(next((k for k,v in KEY_MAP.items() if v == code), "None"))
                
                m_btn = d.get("mouse_btn", 0)
                m_scr = d.get("mouse_scroll", 0)
                self.cb_mouse_btn.set(next((k for k,v in MOUSE_BUTTONS.items() if v == m_btn), "None"))
                self.cb_mouse_scroll.set(next((k for k,v in MOUSE_WHEEL.items() if v == m_scr), "None"))
            
            self.cb_led.set(next((k for k,v in LED_MODES.items() if v == self.led_mode), "Static"))
        except: pass

    def store_ui_state(self, _=None):
        if not self.running or self.is_uploading: return
        tab = self.editor_frame.get()
        idx = self.selected_key_index
        
        if tab == "Input / Macro":
            mod = (1 if self.var_ctrl.get() else 0) | (2 if self.var_shift.get() else 0) | (4 if self.var_alt.get() else 0) | (8 if self.var_win.get() else 0)
            key_code = KEY_MAP.get(self.cb_key.get(), 0)
            mouse_btn = MOUSE_BUTTONS.get(self.cb_mouse_btn.get(), 0)
            mouse_scroll = MOUSE_WHEEL.get(self.cb_mouse_scroll.get(), 0)
            
            if mouse_btn != 0 or mouse_scroll != 0:
                self.current_data[idx] = {
                    "type": "mouse", "mod": mod, "code": 0,
                    "mouse_btn": mouse_btn, "mouse_scroll": mouse_scroll
                }
            else:
                self.current_data[idx] = {
                    "type": "key", "mod": mod, "code": key_code,
                    "mouse_btn": 0, "mouse_scroll": 0
                }
                
        elif tab == "Media":
            b1, b2 = MEDIA_MAP.get(self.cb_media.get(), (0,0))
            self.current_data[idx] = {"type": "media", "b1": b1, "b2": b2}

    def store_led_state(self, _=None):
        self.led_mode = LED_MODES.get(self.cb_led.get(), 1)

    def add_preset(self):
        name = ctk.CTkInputDialog(text="Preset Name:", title="Save").get_input()
        if not name: return
        color = colorchooser.askcolor()[1] or "#888888"
        self.presets[name] = {"keys": [dict(x) for x in self.current_data], "led": self.led_mode, "color": color}
        self.save_presets_file()
        self.refresh_preset_list()
        self.load_preset_by_name(name)

    def del_preset(self):
        if self.current_preset_name in self.presets:
            if self.current_preset_name == self.default_preset_name:
                self.default_preset_name = None
            del self.presets[self.current_preset_name]
            self.save_presets_file()
            self.current_preset_name = None
            self.refresh_preset_list()
            if self.presets:
                self.load_preset_by_name(list(self.presets.keys())[0])

    def start_upload(self):
        if self.is_uploading: return
        if self.pad.is_connected():
            self.set_blocking_state(True)
            threading.Thread(target=self._upload_thread, daemon=True).start()

    def _upload_thread(self):
        with self.upload_lock:
            try:
                self.pad.select_layer(0)
                time.sleep(0.05)
                for i, d in enumerate(self.current_data):
                    t = d.get("type")
                    if t == "key": 
                        self.pad.set_key(i, d["mod"], d["code"])
                    elif t == "media": 
                        self.pad.set_media(i, d["b1"], d["b2"])
                    elif t == "mouse": 
                        self.pad.set_mouse(i, d["mouse_btn"], d["mouse_scroll"], d.get("mod", 0))
                    time.sleep(0.02)
                self.pad.set_led(self.led_mode)
                self.pad.save_to_flash()
            except Exception as e:
                print(f"Upload Error: {e}")
                self.after(0, self.upload_finished, False)
            else:
                self.after(0, self.upload_finished, True)

    def upload_finished(self, success):
        self.set_blocking_state(False)
        if not success:
            if self.running and self.winfo_exists():
                if not self.last_auto_uploaded_preset: 
                    messagebox.showerror("Error", "Upload failed.")

    def set_blocking_state(self, b):
        if not self.running or not self.winfo_exists(): return
        self.is_uploading = b
        s = "disabled" if b else "normal"
        try:
            self.btn_upload.configure(state=s, text="UPLOADING..." if b else "UPLOAD CONFIGURATION")
            self.btn_add.configure(state=s)
            self.btn_del.configure(state=s)
            for btn in self.preset_widgets.values(): btn.configure(state=s)
        except: pass

    def check_conn_loop(self):
        if not self.running: return
        p = self.pad.scan_for_device()
        if p and not self.pad.is_connected():
            self.pad.connect()
        elif not p and self.pad.is_connected():
            self.pad.device = None
            self.pad._connected = False
        
        if self.pad.is_connected() != self.connected_last_frame:
            self.connected_last_frame = self.pad.is_connected()
            self.after(0, self.safe_update_status, self.connected_last_frame)
        
        self.after(2000, self.check_conn_loop)

    def update_status_ui(self, c):
        if not self.running or not self.winfo_exists(): return
        self.update_tray_icon() 
        try:
            self.lbl_status_icon.configure(text_color=Theme.CONNECTED_COLOR if c else Theme.DISCONNECTED_COLOR)
            self.lbl_status_text.configure(text="CONNECTED" if c else "DISCONNECTED", 
                                           text_color=Theme.TEXT_PRIMARY if c else Theme.TEXT_SECONDARY)
            self.btn_upload.configure(state="normal" if c else "disabled", 
                                      fg_color=Theme.CONNECTED_COLOR if c else Theme.WIDGET_BG, 
                                      text_color="black" if c else Theme.TEXT_DISABLED)
            
            if self.init_complete and self.cfg_notify_status:
                self.notify_user("Device Status", "Macropad Connected" if c else "Macropad Disconnected")

            if c and self.current_preset_name:
                self.start_upload()
        except: pass

    def setup_tray(self):
        if self.cfg_tray_enabled and not self.tray_icon:
            self.update_tray_icon()
    
    def update_tray_icon(self):
        if not self.running or not self.cfg_tray_enabled: return
        
        color = "#888888"
        if self.pad.is_connected() and self.current_preset_name in self.presets:
            color = self.presets[self.current_preset_name].get("color", "#888888")
        
        img = Image.new('RGBA', (64, 64), (0,0,0,0))
        d = ImageDraw.Draw(img)
        
        if not self.pad.is_connected():
            d.rounded_rectangle([2, 2, 62, 62], 16, fill="black", outline="#ff3d00", width=4)
            d.line((18, 18, 46, 46), fill="#ff3d00", width=6)
            d.line((46, 18, 18, 46), fill="#ff3d00", width=6)
        else:
            d.rounded_rectangle([2, 2, 62, 62], 16, fill=color, outline=color, width=0)
            d.line((16, 20, 32, 48), fill="white", width=6)
            d.line((32, 48, 48, 20), fill="white", width=6)
        
        if self.tray_icon:
            self.tray_icon.icon = img
            self.tray_icon.menu = self.create_tray_menu()
        else:
            self.tray_icon = pystray.Icon("VMacropad", img, "V Macropad", self.create_tray_menu())
            threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _make_tray_action(self, name):
        # Returns a function with signature (icon, item)
        return lambda icon, item: self.tray_activate_preset(name)

    def _make_tray_check(self, name):
        # Returns a function with signature (item)
        return lambda item: self.current_preset_name == name

    def create_tray_menu(self):
        items = [pystray.MenuItem("Open", self.show_window_tray, default=True), pystray.Menu.SEPARATOR]
        for name in self.presets:
            items.append(pystray.MenuItem(
                name, 
                self._make_tray_action(name), 
                checked=self._make_tray_check(name)
            ))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Quit", self.quit_app))
        return pystray.Menu(*items)

    def tray_activate_preset(self, name):
        if not self.running or self.is_uploading: return
        self.last_auto_uploaded_preset = name
        self.after(0, self.safe_tray_load, name)

    def safe_tray_load(self, name):
        if self.running and self.winfo_exists():
            self.load_preset_by_name(name)
            self.start_upload()

    def show_window_tray(self, icon=None, item=None):
        if self.running:
            self.after(0, self.deiconify)

    def on_close_attempt(self):
        if self.cfg_tray_enabled:
            self.withdraw()
        else:
            self.quit_app()

    def quit_app(self, icon=None, item=None):
        self.running = False
        self.after(0, self._perform_shutdown)

    def _perform_shutdown(self):
        if self.tray_icon:
            self.tray_icon.stop()
        self.destroy()
        os._exit(0)

if __name__ == "__main__":
    app = VMacroApp()
    app.mainloop()
