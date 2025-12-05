import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import json
import shutil
from pathlib import Path
import subprocess
import sys

# Version info - Update this when making releases
VERSION = "2.0.0"
GITHUB_REPO = "Angel2mp3/SheetDL"
GITHUB_RAW_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/SheetDL.py"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/commits/main"

# Check and install missing dependencies
def check_dependencies():
    """Check for required packages and offer to install missing ones"""
    required_packages = {
        'gspread': 'gspread',
        'google.oauth2.service_account': 'google-auth',
        'requests': 'requests',
        'yt_dlp': 'yt-dlp',
        'bs4': 'beautifulsoup4',
        'cloudscraper': 'cloudscraper',
    }
    
    missing = []
    for module, package in required_packages.items():
        try:
            __import__(module.split('.')[0])
        except ImportError:
            missing.append(package)
    
    if missing:
        # Show a simple dialog before the main app loads
        root = tk.Tk()
        root.withdraw()
        
        msg = f"The following packages are required but not installed:\n\n• {chr(10).join(missing)}\n\nWould you like to install them now?"
        if messagebox.askyesno("Missing Dependencies", msg):
            root.destroy()
            
            # Install missing packages
            for package in missing:
                print(f"Installing {package}...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            
            # Restart the script
            print("\nDependencies installed! Restarting...")
            os.execv(sys.executable, [sys.executable] + sys.argv)
        else:
            messagebox.showerror("Cannot Continue", "Required dependencies are missing. The application will now exit.")
            root.destroy()
            sys.exit(1)

check_dependencies()

def check_for_updates(silent=False):
    """Check GitHub for updates and offer to install them"""
    import requests  # Import here since it's after dependency check
    
    try:
        # Get the latest version from GitHub raw file
        response = requests.get(GITHUB_RAW_URL, timeout=10)
        if response.status_code != 200:
            if not silent:
                messagebox.showinfo("Update Check", "Could not check for updates. GitHub may be unavailable.")
            return False
        
        remote_content = response.text
        
        # Extract version from remote file
        remote_version = None
        for line in remote_content.split('\n')[:20]:  # Check first 20 lines
            if line.startswith('VERSION = '):
                # Extract version string
                remote_version = line.split('=')[1].strip().strip('"\'')
                break
        
        if not remote_version:
            if not silent:
                messagebox.showinfo("Update Check", "Could not determine remote version.")
            return False
        
        # Compare versions (simple string comparison, works for semantic versioning)
        def parse_version(v):
            try:
                return tuple(int(x) for x in v.split('.'))
            except:
                return (0, 0, 0)
        
        local_ver = parse_version(VERSION)
        remote_ver = parse_version(remote_version)
        
        if remote_ver > local_ver:
            # New version available
            root = tk.Tk()
            root.withdraw()
            
            msg = f"A new version of SheetDL is available!\n\nCurrent version: {VERSION}\nNew version: {remote_version}\n\nWould you like to update now?"
            if messagebox.askyesno("Update Available", msg):
                root.destroy()
                
                # Get the path of the current script
                current_script = os.path.abspath(sys.argv[0])
                
                # Create backup
                backup_path = current_script + '.backup'
                try:
                    shutil.copy2(current_script, backup_path)
                except Exception as e:
                    messagebox.showerror("Update Failed", f"Could not create backup: {e}")
                    return False
                
                # Write new content
                try:
                    with open(current_script, 'w', encoding='utf-8') as f:
                        f.write(remote_content)
                    
                    # Show success and restart
                    tk.Tk().withdraw()
                    messagebox.showinfo("Update Complete", f"SheetDL has been updated to version {remote_version}!\n\nThe application will now restart.")
                    
                    # Remove backup on success
                    try:
                        os.remove(backup_path)
                    except:
                        pass
                    
                    # Restart the script
                    os.execv(sys.executable, [sys.executable] + sys.argv)
                    
                except Exception as e:
                    # Restore from backup
                    try:
                        shutil.copy2(backup_path, current_script)
                        os.remove(backup_path)
                    except:
                        pass
                    messagebox.showerror("Update Failed", f"Could not write update: {e}\n\nThe original file has been restored.")
                    return False
            else:
                root.destroy()
            return True
        else:
            if not silent:
                tk.Tk().withdraw()
                messagebox.showinfo("No Updates", f"You are running the latest version ({VERSION}).")
            return False
            
    except requests.exceptions.Timeout:
        if not silent:
            messagebox.showinfo("Update Check", "Update check timed out. Please check your internet connection.")
        return False
    except Exception as e:
        if not silent:
            messagebox.showinfo("Update Check", f"Could not check for updates: {e}")
        return False

# Check for updates on startup (silent mode - only prompt if update available)
check_for_updates(silent=True)

# Now import the rest after dependencies are confirmed
import gspread
from google.oauth2.service_account import Credentials
import requests
import zipfile
from urllib.parse import urlparse, parse_qs
import re
import html
from datetime import datetime
import yt_dlp
from bs4 import BeautifulSoup
import ctypes

IS_WINDOWS = os.name == 'nt'
if IS_WINDOWS:
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    shell32 = ctypes.windll.shell32
    kernel32 = ctypes.windll.kernel32
    GWL_EXSTYLE = -20
    WS_EX_APPWINDOW = 0x00040000
    WS_EX_TOOLWINDOW = 0x00000080
    SWP_NOSIZE = 0x0001
    SWP_NOMOVE = 0x0002
    SWP_NOZORDER = 0x0004
    SWP_FRAMECHANGED = 0x0020
    
    # Minimize console window on startup
    def minimize_console():
        try:
            console_hwnd = kernel32.GetConsoleWindow()
            if console_hwnd:
                SW_MINIMIZE = 6
                user32.ShowWindow(console_hwnd, SW_MINIMIZE)
        except Exception:
            pass
    
    minimize_console()
else:
    user32 = None
    gdi32 = None
    shell32 = None
    kernel32 = None


def create_round_rect(canvas, x1, y1, x2, y2, radius=14, **kwargs):
    radius = max(0, min(radius, (x2 - x1) / 2, (y2 - y1) / 2))
    points = [
        x1 + radius, y1,
        x2 - radius, y1,
        x2, y1,
        x2, y1 + radius,
        x2, y2 - radius,
        x2, y2,
        x2 - radius, y2,
        x1 + radius, y2,
        x1, y2,
        x1, y2 - radius,
        x1, y1 + radius,
        x1, y1
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=36, **kwargs)


class RoundedCard(tk.Frame):
    def __init__(self, parent, title, colors, radius=18, padding=18):
        super().__init__(parent, bg=colors["background"])
        self.colors = colors
        self.radius = radius
        self.padding = padding
        self.canvas = tk.Canvas(self, bg=colors["background"], highlightthickness=0, bd=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", self._redraw)

        self.container = tk.Frame(self.canvas, bg=colors["panel"], bd=0, highlightthickness=0)
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(1, weight=1)
        self.window = self.canvas.create_window(self.padding, self.padding, anchor='nw', window=self.container)

        self.title_label = tk.Label(
            self.container,
            text=title,
            bg=colors["panel"],
            fg=colors["text"],
            font=("Segoe UI", 12, "bold")
        )
        self.title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))

        self.body = tk.Frame(self.container, bg=colors["panel"])
        self.body.grid(row=1, column=0, sticky='nsew')

    def _redraw(self, event=None):
        self.canvas.delete("card")
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        if width <= 0 or height <= 0:
            return
        create_round_rect(
            self.canvas,
            0,
            0,
            width,
            height,
            radius=self.radius,
            fill=self.colors["panel"],
            outline=self.colors["panel"],
            tags=("card",)
        )
        inner_width = max(0, width - (self.padding * 2))
        inner_height = max(0, height - (self.padding * 2))
        self.canvas.coords(self.window, self.padding, self.padding)
        self.canvas.itemconfig(self.window, width=inner_width, height=inner_height)


class RoundedButton(tk.Frame):
    def __init__(self, parent, text, command, colors, width=160, height=44):
        # Get the actual background color from the parent widget
        parent_bg = parent.cget('bg') if hasattr(parent, 'cget') else colors["panel"]
        super().__init__(parent, bg=parent_bg, highlightthickness=0, bd=0)
        self.colors = colors
        self.parent_bg = parent_bg
        self.command = command
        self.text = text
        self.width = width
        self.height = height
        self.radius = 20
        self.enabled = True
        self.hover = False

        self.canvas = tk.Canvas(
            self,
            width=self.width,
            height=self.height,
            bg=parent_bg,
            highlightthickness=0,
            bd=0
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Button-1>", self._on_click)
        self._draw()

    def _draw(self):
        self.canvas.delete("all")
        fill = self.colors["accent"]
        text_color = self.colors["text"]
        if not self.enabled:
            fill = "#5b1d31"
            text_color = "#b27f8f"
        elif self.hover:
            fill = self.colors["highlight"]
        create_round_rect(
            self.canvas,
            2,
            2,
            self.width - 2,
            self.height - 2,
            radius=self.radius,
            fill=fill,
            outline=""
        )
        self.canvas.create_text(
            self.width / 2,
            self.height / 2,
            text=self.text,
            fill=text_color,
            font=("Segoe UI", 10, "bold")
        )

    def _on_enter(self, _event=None):
        self.hover = True
        self._draw()

    def _on_leave(self, _event=None):
        self.hover = False
        self._draw()

    def _on_click(self, _event=None):
        if self.enabled and callable(self.command):
            self.command()

    def config(self, **kwargs):
        if 'state' in kwargs:
            self.set_state(kwargs['state'])
        if 'text' in kwargs:
            self.text = kwargs['text']
            self._draw()

    def set_state(self, state):
        self.enabled = state != 'disabled'
        self._draw()
    
    def update_text(self, new_text):
        """Update the button text"""
        self.text = new_text
        self._draw()

    def grid(self, *args, **kwargs):
        super().grid(*args, **kwargs)

    def pack(self, *args, **kwargs):
        super().pack(*args, **kwargs)


class MusicDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"SheetDL v{VERSION}")
        self.root.geometry("900x700")
        self.colors = {
            "background": "#420516",
            "panel": "#7D1935",
            "accent": "#B42B51",
            "highlight": "#E63E6D",
            "text": "#FFE6EE",
            "text_dim": "#F5CAD1"
        }
        self.root.configure(bg=self.colors["background"])
        self.root.minsize(900, 650)
        
        # Set window title and icon BEFORE overrideredirect
        self.root.title(f"SheetDL v{VERSION}")
        self._set_initial_icon()
        
        # Hide window initially, will show after taskbar setup
        self.root.withdraw()
        
        self.root.overrideredirect(True)
        self.is_maximized = False
        self.prev_geometry = None
        self.drag_position = (0, 0)
        self.resizing = False
        self.resize_dir = None
        self.hwnd = None
        self.window_round_radius = 26
        self.log_auto_follow = True
        self.cover_cache = set()
        self.icon_image = None
        self.sheet_tabs = []
        self.selected_tab_gid = None
        self.default_headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        self.setup_title_bar()
        
        # Configuration
        self.config_file = "config.json"
        self.load_config()
        
        # Variables
        self.is_downloading = False
        self.is_paused = False
        self.download_thread = None
        self.working_csv_url = None  # Store the working CSV URL
        self.current_sheet_name = "Sheet"
        self.sheet_tab_var = tk.StringVar(value="Loading…")
        
        # Download queue
        self.download_queue = []  # List of {'url': str, 'gid': str, 'name': str}
        
        self.setup_ui()
        self.create_resize_handles()
        
        # Initialize window after a delay to ensure it's fully created
        self.root.after(100, self._initialize_window)
        
    def load_config(self):
        """Load saved configuration"""
        default_config = {
            "sheet_url": "",
            "gid": "0",
            "output_folder": str(Path.home() / "Downloads" / "Music"),
            "organize_by": "artist",
            "create_zip": False,
            "save_metadata": True,
            "column_mapping": {
                "artist": "Artist",
                "title": "Title",
                "url": "AUTO",
                "genre": "Genre",
                "cover": ""
            }
        }
        
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                loaded = json.load(f)
        else:
            loaded = {}
        
        # Merge defaults with saved values so new settings always exist
        self.config = {**default_config, **loaded}
        saved_mapping = loaded.get("column_mapping", {}) if 'loaded' in locals() else {}
        self.config["column_mapping"] = {
            **default_config["column_mapping"],
            **self.config.get("column_mapping", {}),
            **saved_mapping
        }
            
    def save_config(self):
        """Save current configuration"""
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=2)
    
    def setup_title_bar(self):
        self.title_bar = tk.Frame(self.root, bg=self.colors["panel"], relief='flat', bd=0, height=40)
        self.title_bar.pack(fill=tk.X, side=tk.TOP)
        self.title_bar.pack_propagate(False)

        self.title_label = tk.Label(
            self.title_bar,
            text=f"SheetDL v{VERSION}",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold")
        )
        self.title_label.pack(side=tk.LEFT, padx=(15, 5))

        # Clickable rainbow author credit with rounded gray background
        self._author_text = "By Angel2mp3"
        self._author_font = ("Segoe UI", 9, "bold")
        
        # Calculate text width for background sizing
        self.author_canvas = tk.Canvas(
            self.title_bar,
            bg=self.colors["panel"],
            highlightthickness=0,
            height=24,
            width=95,
            cursor="hand2"
        )
        self.author_canvas.pack(side=tk.LEFT, pady=8)
        self.author_canvas.bind("<Button-1>", lambda e: self._open_github())
        
        # Draw rounded rectangle background
        self._author_bg = create_round_rect(
            self.author_canvas, 0, 0, 95, 24,
            radius=8, fill="#3A0A17", outline="#310815"
        )
        
        # Create individual text items for each character
        self._author_labels = []
        self._gradient_colors = self._generate_gradient_colors()
        
        x_pos = 8
        for i, char in enumerate(self._author_text):
            text_id = self.author_canvas.create_text(
                x_pos, 12,
                text=char,
                fill=self._gradient_colors[i % len(self._gradient_colors)],
                font=self._author_font,
                anchor="w"
            )
            self._author_labels.append(text_id)
            # Measure char width - tighter for no spacing
            if char == " ":
                x_pos += 4
            elif char in "lij1":
                x_pos += 5
            elif char in "mwMW":
                x_pos += 9
            else:
                x_pos += 7
        
        self._author_color_offset = 0.0
        self._animate_author_rainbow()

        control_frame = tk.Frame(self.title_bar, bg=self.colors["panel"])
        control_frame.pack(side=tk.RIGHT, padx=5)

        self.close_btn = self._create_title_button(control_frame, "✕", self.close_window)
        self.max_btn = self._create_title_button(control_frame, "□", self.toggle_max_restore)
        self.min_btn = self._create_title_button(control_frame, "–", self.minimize_window)

        for widget in (self.title_bar, self.title_label):
            widget.bind("<ButtonPress-1>", self.start_move)
            widget.bind("<B1-Motion>", self.do_move)
            widget.bind("<Double-Button-1>", lambda _e: self.toggle_max_restore())

        # Subtle shadow line under title bar
        self.title_shadow = tk.Frame(self.root, bg="#1a0008", height=2)
        self.title_shadow.pack(fill=tk.X, side=tk.TOP)

        self.body_frame = tk.Frame(self.root, bg=self.colors["background"], bd=0, highlightthickness=0)
        self.body_frame.pack(fill=tk.BOTH, expand=True)

    def _generate_gradient_colors(self):
        """Generate smooth gradient colors for rainbow effect"""
        base_colors = [
            (255, 107, 157),  # Pink
            (255, 140, 105),  # Orange
            (255, 217, 61),   # Yellow
            (107, 203, 119),  # Green
            (77, 150, 255),   # Blue
            (155, 89, 182),   # Purple
            (232, 67, 147),   # Magenta
        ]
        
        # Create more colors by interpolating between base colors
        gradient = []
        steps_between = 8
        for i in range(len(base_colors)):
            c1 = base_colors[i]
            c2 = base_colors[(i + 1) % len(base_colors)]
            for step in range(steps_between):
                t = step / steps_between
                r = int(c1[0] + (c2[0] - c1[0]) * t)
                g = int(c1[1] + (c2[1] - c1[1]) * t)
                b = int(c1[2] + (c2[2] - c1[2]) * t)
                gradient.append(f"#{r:02x}{g:02x}{b:02x}")
        return gradient

    def _animate_author_rainbow(self):
        """Animate smooth rainbow wave across letters - left to right"""
        try:
            num_colors = len(self._gradient_colors)
            for i, text_id in enumerate(self._author_labels):
                # Spread colors across text, offset shifts left-to-right
                color_idx = int((i * 4 - self._author_color_offset) % num_colors)
                self.author_canvas.itemconfig(text_id, fill=self._gradient_colors[color_idx])
            self._author_color_offset = (self._author_color_offset + 1) % num_colors
            self.root.after(80, self._animate_author_rainbow)  # Smooth but slow wave
        except tk.TclError:
            pass  # Window closed

    def _open_github(self):
        """Open GitHub profile in browser"""
        import webbrowser
        webbrowser.open("https://github.com/Angel2mp3")

    def _create_title_button(self, parent, symbol, command):
        btn = tk.Button(
            parent,
            text=symbol,
            command=command,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            activebackground=self.colors["highlight"],
            activeforeground=self.colors["text"],
            borderwidth=0,
            relief='flat',
            width=4,
            font=("Segoe UI", 10, "bold")
        )
        btn.pack(side=tk.RIGHT, padx=2, pady=5)
        return btn

    def _set_initial_icon(self):
        """Set icon before overrideredirect for proper taskbar icon"""
        icon_path = Path(__file__).parent / 'app.ico'
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

    def minimize_window(self):
        """Minimize window to taskbar properly"""
        self._is_minimized = True
        if IS_WINDOWS and self.hwnd:
            try:
                SW_MINIMIZE = 6
                user32.ShowWindow(self.hwnd, SW_MINIMIZE)
            except Exception:
                self.root.iconify()
        else:
            self.root.iconify()
    
    def restore_window(self):
        """Restore window from minimized state"""
        self._is_minimized = False
        if IS_WINDOWS and self.hwnd:
            try:
                SW_RESTORE = 9
                user32.ShowWindow(self.hwnd, SW_RESTORE)
                self.root.lift()
                self.root.focus_force()
            except Exception:
                self.root.deiconify()
        else:
            self.root.deiconify()
    
    def _on_focus_in(self, event=None):
        """Handle focus gained (e.g., from taskbar click)"""
        if hasattr(self, '_is_minimized') and self._is_minimized:
            self.restore_window()
    
    def _on_visibility_change(self, event=None):
        """Handle visibility changes"""
        pass
    
    def _check_window_state(self):
        """Periodically check if window was clicked in taskbar"""
        if not IS_WINDOWS or not self.hwnd:
            return
        try:
            # Check if window is minimized
            GWL_STYLE = -16
            WS_MINIMIZE = 0x20000000
            style = user32.GetWindowLongW(self.hwnd, GWL_STYLE)
            is_minimized = bool(style & WS_MINIMIZE)
            
            if hasattr(self, '_is_minimized'):
                if self._is_minimized and not is_minimized:
                    # Window was restored (taskbar click)
                    self._is_minimized = False
                    self.root.lift()
                    self.root.focus_force()
                elif not self._is_minimized and is_minimized:
                    self._is_minimized = True
        except Exception:
            pass
        
        # Check again after 100ms
        self.root.after(100, self._check_window_state)

    def toggle_max_restore(self):
        if not self.is_maximized:
            self.prev_geometry = self.root.geometry()
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            self.root.geometry(f"{screen_w}x{screen_h}+0+0")
            self.is_maximized = True
            self.max_btn.config(text="❐")
        else:
            if self.prev_geometry:
                self.root.geometry(self.prev_geometry)
            self.is_maximized = False
            self.max_btn.config(text="□")
        self.apply_window_rounding()

    def close_window(self):
        self.root.destroy()

    def start_move(self, event):
        self.drag_position = (event.x_root, event.y_root)

    def do_move(self, event):
        if self.is_maximized:
            return
        dx = event.x_root - self.drag_position[0]
        dy = event.y_root - self.drag_position[1]
        x = self.root.winfo_x() + dx
        y = self.root.winfo_y() + dy
        self.root.geometry(f"+{x}+{y}")
        self.drag_position = (event.x_root, event.y_root)

    def _restore_overrideredirect(self, _event=None):
        """Restore overrideredirect after minimize"""
        pass  # No longer needed with new approach

    def _initialize_window(self):
        """Initialize window handles and taskbar after window is ready"""
        try:
            # Make window visible first so we can get the handle
            self.root.deiconify()
            self.root.update_idletasks()
            
            if not IS_WINDOWS:
                self.root.lift()
                self.root.focus_force()
                return
            
            # Get the actual top-level window handle
            GW_OWNER = 4
            hwnd = self.root.winfo_id()
            parent = user32.GetParent(hwnd)
            if parent:
                hwnd = parent
            self.hwnd = hwnd
            
            # Set app ID for taskbar grouping
            if shell32:
                shell32.SetCurrentProcessExplicitAppUserModelID("SheetDL.App")
            
            # Make window appear in taskbar
            style = user32.GetWindowLongW(self.hwnd, GWL_EXSTYLE)
            style &= ~WS_EX_TOOLWINDOW
            style |= WS_EX_APPWINDOW
            user32.SetWindowLongW(self.hwnd, GWL_EXSTYLE, style)
            
            # Set the icon
            icon_path = Path(__file__).parent / 'app.ico'
            if icon_path.exists():
                WM_SETICON = 0x0080
                ICON_BIG = 1
                ICON_SMALL = 0
                LR_LOADFROMFILE = 0x0010
                LR_DEFAULTSIZE = 0x0040
                
                hicon = user32.LoadImageW(
                    None,
                    str(icon_path),
                    1,  # IMAGE_ICON
                    0, 0,
                    LR_LOADFROMFILE | LR_DEFAULTSIZE
                )
                if hicon:
                    user32.SendMessageW(self.hwnd, WM_SETICON, ICON_BIG, hicon)
                    user32.SendMessageW(self.hwnd, WM_SETICON, ICON_SMALL, hicon)
            
            # Refresh window to apply taskbar changes
            user32.SetWindowPos(
                self.hwnd, 0, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
            )
            
            # Apply window rounding
            self.apply_window_rounding()
            
            # Bind to window state changes for taskbar click handling
            self.root.bind('<FocusIn>', self._on_focus_in)
            self.root.bind('<Visibility>', self._on_visibility_change)
            self._is_minimized = False
            
            # Bring to front
            self.root.lift()
            self.root.focus_force()
            
            # Start checking for window activation (for taskbar clicks)
            self._check_window_state()
            
        except Exception as e:
            print(f"Window init error: {e}")
            # Make sure window is visible even if there's an error
            self.root.deiconify()

    def _flash_topmost(self):
        try:
            self.root.attributes('-topmost', True)
            self.root.after(200, lambda: self.root.attributes('-topmost', False))
        except tk.TclError:
            pass

    def apply_window_rounding(self):
        if not IS_WINDOWS or not self.hwnd or not gdi32:
            return
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        if width <= 0 or height <= 0:
            return
        radius = 0 if self.is_maximized else self.window_round_radius
        region = gdi32.CreateRoundRectRgn(0, 0, width, height, radius, radius)
        user32.SetWindowRgn(self.hwnd, region, True)

    def _on_window_configure(self, _event=None):
        self.apply_window_rounding()

    def create_resize_handles(self):
        border = 6
        specs = {
            'e': dict(relx=1.0, rely=0, anchor='ne', width=border, relheight=1.0, cursor='sb_h_double_arrow'),
            's': dict(relx=0, rely=1.0, anchor='sw', height=border, relwidth=1.0, cursor='sb_v_double_arrow'),
            'se': dict(relx=1.0, rely=1.0, anchor='se', width=border, height=border, cursor='size_nw_se')
        }
        self.resize_handles = {}
        for direction, spec in specs.items():
            handle = tk.Frame(self.root, bg=self.colors["background"], cursor=spec['cursor'])
            handle.place(**{k: v for k, v in spec.items() if k not in ('cursor',)})
            handle.bind("<ButtonPress-1>", lambda e, d=direction: self.start_resize(e, d))
            handle.bind("<B1-Motion>", self.perform_resize)
            handle.bind("<ButtonRelease-1>", self.stop_resize)
            self.resize_handles[direction] = handle

    def start_resize(self, event, direction):
        self.resizing = True
        self.resize_dir = direction
        self.resize_start = (event.x_root, event.y_root)
        self.start_width = self.root.winfo_width()
        self.start_height = self.root.winfo_height()

    def perform_resize(self, event):
        if not self.resizing:
            return
        dx = event.x_root - self.resize_start[0]
        dy = event.y_root - self.resize_start[1]
        new_width = self.start_width
        new_height = self.start_height
        if 'e' in self.resize_dir:
            new_width += dx
        if 's' in self.resize_dir:
            new_height += dy
        min_w = self.root.minsize()[0]
        min_h = self.root.minsize()[1]
        new_width = max(min_w, new_width)
        new_height = max(min_h, new_height)
        self.root.geometry(f"{new_width}x{new_height}+{self.root.winfo_x()}+{self.root.winfo_y()}")
        self.apply_window_rounding()

    def stop_resize(self, _event=None):
        self.resizing = False
        self.resize_dir = None

    def setup_ui(self):
        """Setup the user interface"""
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', background=self.colors["panel"], foreground=self.colors["text"])

        style.configure('.', background=self.colors["panel"], foreground=self.colors["text"])
        style.configure('Main.TFrame', background=self.colors["background"])
        style.configure('Card.TLabelframe', background=self.colors["panel"], foreground=self.colors["text"], borderwidth=0, relief='flat')
        style.configure('Card.TLabelframe.Label', background=self.colors["panel"], foreground=self.colors["text"], font=('Segoe UI', 11, 'bold'))
        style.configure('TLabel', background=self.colors["panel"], foreground=self.colors["text"], font=('Segoe UI', 10))
        style.configure('Primary.TButton', background=self.colors["accent"], foreground=self.colors["text"], padding=(20, 8), relief='flat')
        style.map('Primary.TButton', background=[('active', self.colors["highlight"])])
        style.configure('Link.TButton', background=self.colors["panel"], foreground=self.colors["highlight"], padding=(10, 4), relief='flat')
        style.map('Link.TButton', foreground=[('active', self.colors["text"])] )
        
        # Entry styling with consistent colors on focus
        style.configure('TEntry', 
            fieldbackground='#4F0F24', 
            foreground=self.colors["text"], 
            insertcolor=self.colors["text"],
            selectbackground=self.colors["accent"],
            selectforeground=self.colors["text"]
        )
        style.map('TEntry',
            fieldbackground=[('focus', '#5A1228'), ('!focus', '#4F0F24')],
            selectbackground=[('focus', self.colors["accent"])],
            selectforeground=[('focus', self.colors["text"])]
        )
        
        # Combobox: dark background with light text
        style.configure('TCombobox', 
            fieldbackground='#4F0F24', 
            background=self.colors["panel"], 
            foreground=self.colors["text"],
            selectbackground=self.colors["panel"],
            selectforeground=self.colors["text"],
            arrowcolor=self.colors["text"]
        )
        style.map('TCombobox',
            fieldbackground=[('readonly', '#4F0F24'), ('disabled', '#3a0a1a'), ('focus', '#4F0F24')],
            foreground=[('readonly', self.colors["text"]), ('disabled', self.colors["text_dim"])],
            background=[('readonly', self.colors["panel"])],
            selectbackground=[('readonly', '#4F0F24'), ('focus', '#4F0F24')],
            selectforeground=[('readonly', self.colors["text"]), ('focus', self.colors["text"])]
        )
        # Fix combobox dropdown list colors
        self.root.option_add('*TCombobox*Listbox.background', '#4F0F24')
        self.root.option_add('*TCombobox*Listbox.foreground', self.colors["text"])
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.colors["accent"])
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.colors["text"])
        
        # Scrollbar styling - dark theme
        style.configure('Vertical.TScrollbar',
            background=self.colors["panel"],
            troughcolor=self.colors["background"],
            bordercolor=self.colors["background"],
            arrowcolor=self.colors["text"],
            relief='flat'
        )
        style.map('Vertical.TScrollbar',
            background=[('active', self.colors["accent"]), ('pressed', self.colors["highlight"])]
        )
        # Checkbox: matching colors with proper hover states
        style.configure('TCheckbutton', 
            background=self.colors["panel"], 
            foreground=self.colors["text"],
            focuscolor=self.colors["panel"]
        )
        style.map('TCheckbutton',
            background=[('active', self.colors["panel"]), ('pressed', self.colors["panel"])],
            foreground=[('active', self.colors["highlight"]), ('pressed', self.colors["text"])],
            indicatorcolor=[('selected', self.colors["highlight"]), ('pressed', self.colors["accent"])]
        )
        style.configure('Accent.Horizontal.TProgressbar', background=self.colors["highlight"], troughcolor='#4F0F24', thickness=18)
        
        # Scrollable content area
        self.canvas = tk.Canvas(self.body_frame, bg=self.colors["background"], highlightthickness=0, bd=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = ttk.Scrollbar(self.body_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.content_frame = tk.Frame(self.canvas, bg=self.colors["background"], padx=20, pady=20)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor='nw')
        self.content_frame.bind("<Configure>", self._update_scrollregion)
        self.canvas.bind("<Configure>", self._sync_canvas_width)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add='+')

        # Main container with padding
        main_frame = self.content_frame
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Google Sheets Section
        sheets_section = RoundedCard(main_frame, "Google Sheets Settings", self.colors)
        sheets_section.grid(row=0, column=0, sticky='ew')
        sheets_frame = sheets_section.body
        sheets_frame.columnconfigure(1, weight=1)
        sheets_frame.columnconfigure(2, weight=1)
        
        ttk.Label(sheets_frame, text="Sheet URL:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.sheet_url_var = tk.StringVar(value=self.config.get("sheet_url", ""))
        sheet_url_entry = ttk.Entry(sheets_frame, textvariable=self.sheet_url_var, width=60)
        sheet_url_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=(0,5))
        
        # Button frame for Test Connection and Add to Queue
        sheet_btn_frame = tk.Frame(sheets_frame, bg=self.colors["panel"])
        sheet_btn_frame.grid(row=0, column=3, padx=(10,0), pady=(0,5))
        
        self.test_connection_btn = RoundedButton(sheet_btn_frame, "Test Connection", self.test_connection, self.colors, width=140, height=38)
        self.test_connection_btn.pack(side=tk.TOP, pady=(0, 2))
        
        self.queue_btn = RoundedButton(sheet_btn_frame, "Add to Queue", self.add_to_queue, self.colors, width=140, height=38)
        self.queue_btn.pack(side=tk.TOP)
        self.queue_btn.set_state('disabled')
        
        ttk.Label(sheets_frame, text="Tab/Sheet GID (optional):").grid(row=1, column=0, sticky=tk.W)
        self.gid_var = tk.StringVar(value=self.config.get("gid", "0"))
        gid_entry = ttk.Entry(sheets_frame, textvariable=self.gid_var, width=20)
        gid_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(sheets_frame, text="(Leave as 0 for first tab, or extract from URL: gid=123)").grid(
            row=1, column=2, columnspan=2, sticky=tk.W, padx=5
        )
        ttk.Label(sheets_frame, text="Sheet Tab:").grid(row=2, column=0, sticky=tk.W, pady=(5,0))
        self.sheet_tab_combo = ttk.Combobox(
            sheets_frame,
            textvariable=self.sheet_tab_var,
            state='disabled',
            width=30
        )
        self.sheet_tab_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=(5,0))
        self.sheet_tab_combo.bind("<<ComboboxSelected>>", self.on_tab_selected)
        self.tab_hint = tk.Label(
            sheets_frame,
            text="Connect to load sheet tabs",
            bg=self.colors["panel"],
            fg=self.colors["text_dim"],
            font=("Segoe UI", 9)
        )
        self.tab_hint.grid(row=2, column=2, columnspan=2, sticky=tk.W, pady=(5,0))
        
        # Column Mapping Section
        mapping_section = RoundedCard(main_frame, "Column Mapping", self.colors)
        mapping_section.grid(row=1, column=0, sticky='ew', pady=(15, 0))
        mapping_frame = mapping_section.body
        
        ttk.Label(mapping_frame, text="Artist Column:").grid(row=0, column=0, sticky=tk.W)
        self.artist_col_var = tk.StringVar(value=self.config["column_mapping"].get("artist", "Artist"))
        ttk.Entry(mapping_frame, textvariable=self.artist_col_var, width=20).grid(row=0, column=1, padx=5)
        
        ttk.Label(mapping_frame, text="Title Column:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.title_col_var = tk.StringVar(value=self.config["column_mapping"].get("title", "Title"))
        ttk.Entry(mapping_frame, textvariable=self.title_col_var, width=20).grid(row=0, column=3, padx=5)
        
        ttk.Label(mapping_frame, text="URL Column:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.url_col_var = tk.StringVar(value=self.config["column_mapping"].get("url", "AUTO"))
        ttk.Entry(mapping_frame, textvariable=self.url_col_var, width=20).grid(row=1, column=1, padx=5, pady=(5, 0))
        
        ttk.Label(mapping_frame, text="Genre Column (optional):").grid(row=1, column=2, sticky=tk.W, padx=(20, 0), pady=(5, 0))
        self.genre_col_var = tk.StringVar(value=self.config["column_mapping"].get("genre", "Genre"))
        ttk.Entry(mapping_frame, textvariable=self.genre_col_var, width=20).grid(row=1, column=3, padx=5, pady=(5, 0))
        ttk.Label(mapping_frame, text="Cover/Image Column (optional):").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        self.cover_col_var = tk.StringVar(value=self.config["column_mapping"].get("cover", ""))
        ttk.Entry(mapping_frame, textvariable=self.cover_col_var, width=20).grid(row=2, column=1, padx=5, pady=(5, 0))
        
        ttk.Label(mapping_frame, text="(Leave 'AUTO' to auto-detect link column)").grid(
            row=3,
            column=0,
            columnspan=3,
            sticky=tk.W,
            pady=(2, 0)
        )
        
        # Output Settings Section
        output_section = RoundedCard(main_frame, "Output Settings", self.colors)
        output_section.grid(row=2, column=0, sticky='ew', pady=(15, 0))
        output_frame = output_section.body
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Folder:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        self.output_folder_var = tk.StringVar(value=self.config.get("output_folder", ""))
        ttk.Entry(output_frame, textvariable=self.output_folder_var, width=50).grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=5
        )
        self.browse_btn = RoundedButton(output_frame, "Browse", self.browse_folder, self.colors, width=120, height=38)
        self.browse_btn.grid(row=0, column=2, padx=5)
        
        ttk.Label(output_frame, text="Organize By:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.organize_var = tk.StringVar(value=self.config.get("organize_by", "artist"))
        organize_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.organize_var,
            values=["none", "artist", "genre", "artist_genre"],
            state="readonly",
            width=20
        )
        organize_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(5, 0))
        
        self.create_zip_var = tk.BooleanVar(value=self.config.get("create_zip", False))
        ttk.Checkbutton(
            output_frame,
            text="Create ZIP archive after download",
            variable=self.create_zip_var
        ).grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        self.save_log_var = tk.BooleanVar(value=self.config.get("save_log", False))
        ttk.Checkbutton(
            output_frame,
            text="Save download log to text file",
            variable=self.save_log_var
        ).grid(row=2, column=1, sticky=tk.W, padx=(30, 0), pady=(5, 0))
        
        self.save_metadata_var = tk.BooleanVar(value=self.config.get("save_metadata", True))
        ttk.Checkbutton(
            output_frame,
            text="Save metadata text files for each song",
            variable=self.save_metadata_var
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # YouTube Format Selection
        ttk.Label(output_frame, text="YouTube Format:").grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
        self.yt_format_var = tk.StringVar(value=self.config.get("yt_format", "video_mp4"))
        yt_format_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.yt_format_var,
            values=["video_mp4", "video_best", "audio_m4a", "audio_mp3"],
            state="readonly",
            width=15
        )
        yt_format_combo.grid(row=4, column=1, sticky=tk.W, padx=5, pady=(10, 0))
        
        # SoundCloud Format Selection
        ttk.Label(output_frame, text="SoundCloud Format:").grid(row=5, column=0, sticky=tk.W, pady=(5, 0))
        self.sc_format_var = tk.StringVar(value=self.config.get("sc_format", "audio_m4a"))
        sc_format_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.sc_format_var,
            values=["audio_m4a", "audio_mp3"],
            state="readonly",
            width=15
        )
        sc_format_combo.grid(row=5, column=1, sticky=tk.W, padx=5, pady=(5, 0))
        
        # Progress Section
        progress_section = RoundedCard(main_frame, "Download Progress", self.colors)
        progress_section.grid(row=3, column=0, sticky='nsew', pady=(15, 0))
        main_frame.rowconfigure(3, weight=1)
        progress_frame = progress_section.body
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate',
            style='Accent.Horizontal.TProgressbar'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=15, width=80, state='disabled')
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.configure(
            bg=self.colors["background"],
            fg=self.colors["text_dim"],
            insertbackground=self.colors["highlight"],
            font=("Consolas", 10),
            relief='flat',
            borderwidth=0
        )
        for event_name in (
            "<MouseWheel>", "<Shift-MouseWheel>", "<Button-4>", "<Button-5>",
            "<KeyRelease-Up>", "<KeyRelease-Down>", "<KeyRelease-Prior>", "<KeyRelease-Next>",
            "<ButtonRelease-1>"
        ):
            self.log_text.bind(event_name, self.on_log_manual_scroll)
        self.jump_to_latest_btn = RoundedButton(
            progress_frame, 
            "Jump to Latest", 
            self.enable_log_autoscroll, 
            self.colors, 
            width=130, 
            height=32
        )
        self.jump_to_latest_btn.grid(row=2, column=0, sticky=tk.E, pady=(6, 0))
        self.jump_to_latest_btn.set_state('disabled')
        self.refresh_log_follow_state()
        
        # Control Buttons
        button_frame = tk.Frame(main_frame, bg=self.colors["background"])
        button_frame.grid(row=4, column=0, pady=(20, 0))
        
        self.download_btn = RoundedButton(button_frame, "Start Download", self.start_download, self.colors, width=180)
        self.download_btn.grid(row=0, column=0, padx=8)
        
        self.pause_btn = RoundedButton(button_frame, "Pause", self.toggle_pause, self.colors, width=100)
        self.pause_btn.grid(row=0, column=1, padx=8)
        self.pause_btn.set_state('disabled')
        
        self.stop_btn = RoundedButton(button_frame, "Stop", self.stop_download, self.colors, width=100)
        self.stop_btn.grid(row=0, column=2, padx=8)
        self.stop_btn.set_state('disabled')
        
        self.save_btn = RoundedButton(button_frame, "Save Settings", self.save_settings, self.colors, width=160)
        self.save_btn.grid(row=0, column=3, padx=8)
        
        self.update_btn = RoundedButton(button_frame, "Check for Updates", self.check_for_updates_gui, self.colors, width=180)
        self.update_btn.grid(row=0, column=4, padx=8)
        
        # Queue Display (compact)
        self.queue_frame = tk.Frame(main_frame, bg=self.colors["background"])
        self.queue_frame.grid(row=5, column=0, pady=(10, 0), sticky='ew')
        self.queue_frame.grid_remove()  # Hidden by default
        
        queue_header = tk.Frame(self.queue_frame, bg=self.colors["panel"], height=28)
        queue_header.pack(fill=tk.X)
        queue_header.pack_propagate(False)
        
        self.queue_label = tk.Label(
            queue_header,
            text="📋 Queue (0)",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 9, "bold")
        )
        self.queue_label.pack(side=tk.LEFT, padx=10)
        
        self.clear_queue_btn = tk.Label(
            queue_header,
            text="Clear All",
            bg=self.colors["panel"],
            fg=self.colors["highlight"],
            font=("Segoe UI", 9),
            cursor="hand2"
        )
        self.clear_queue_btn.pack(side=tk.RIGHT, padx=10)
        self.clear_queue_btn.bind("<Button-1>", lambda e: self.clear_queue())
        
        self.queue_listbox = tk.Listbox(
            self.queue_frame,
            height=3,
            bg=self.colors["background"],
            fg=self.colors["text_dim"],
            font=("Segoe UI", 9),
            relief='flat',
            borderwidth=0,
            selectbackground=self.colors["accent"],
            selectforeground=self.colors["text"],
            activestyle='none'
        )
        self.queue_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.queue_listbox.bind("<Double-Button-1>", self.remove_queue_item)

        self.report_environment()
        
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        if self.log_auto_follow:
            self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()
        self.refresh_log_follow_state()

    def on_log_manual_scroll(self, _event=None):
        self.root.after_idle(self.refresh_log_follow_state)

    def enable_log_autoscroll(self):
        self.log_auto_follow = True
        self.log_text.see(tk.END)
        self.refresh_log_follow_state()

    def refresh_log_follow_state(self):
        if not hasattr(self, 'log_text'):
            return
        _, last = self.log_text.yview()
        self.log_auto_follow = last >= 0.999
        if hasattr(self, 'jump_to_latest_btn'):
            state = 'disabled' if self.log_auto_follow else 'normal'
            self.jump_to_latest_btn.set_state(state)

    def _update_scrollregion(self, _event=None):
        if hasattr(self, 'canvas'):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_canvas_width(self, event):
        if hasattr(self, 'canvas'):
            self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if not hasattr(self, 'canvas'):
            return
        widget_path = str(event.widget)
        prefix = str(self.content_frame)
        if hasattr(self, 'log_text'):
            log_base = str(self.log_text)
            if widget_path.startswith(log_base) or widget_path.startswith(f"{log_base}."):
                return
        if not widget_path.startswith(prefix):
            return
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        if self.canvas.winfo_height() < (bbox[3] - bbox[1]):
            delta = -1 * int(event.delta / 120)
            self.canvas.yview_scroll(delta, 'units')
            return "break"

    def fetch_sheet_tabs(self, sheet_id):
        try:
            url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
            response = requests.get(url, headers=self.default_headers, timeout=15)
            if response.status_code != 200:
                return []
            text = response.text
            pattern = re.compile(r'"gid":(\d+),"title":"(.*?)"')
            matches = pattern.findall(text)
            tabs = []
            seen = set()
            for gid, raw_title in matches:
                if gid in seen:
                    continue
                seen.add(gid)
                decoded = html.unescape(raw_title)
                tabs.append({"gid": gid, "title": decoded})
            return tabs
        except Exception:
            return []

    def _choose_default_tab(self, tabs):
        priorities = ['unreleased', 'main']
        for keyword in priorities:
            for tab in tabs:
                if keyword in tab['title'].lower():
                    return tab
        return tabs[0]

    def update_sheet_tab_options(self, tabs, preferred_gid):
        if not tabs:
            self.sheet_tab_combo.set("Single tab")
            self.sheet_tab_combo.config(state='disabled')
            self.tab_hint.config(text="No alternate tabs detected")
            chosen_gid = preferred_gid or self.gid_var.get() or "0"
            self.selected_tab_gid = chosen_gid
            self.gid_var.set(chosen_gid)
            return chosen_gid

        self.sheet_tabs = tabs
        self.sheet_tab_combo['values'] = [tab['title'] for tab in tabs]
        selected_tab = None
        if preferred_gid:
            for tab in tabs:
                if tab['gid'] == preferred_gid:
                    selected_tab = tab
                    break
        if not selected_tab:
            selected_tab = self._choose_default_tab(tabs)
        self.sheet_tab_var.set(selected_tab['title'])
        self.selected_tab_gid = selected_tab['gid']
        self.gid_var.set(self.selected_tab_gid)
        combo_state = 'disabled' if len(tabs) == 1 else 'readonly'
        self.sheet_tab_combo.config(state=combo_state)
        summary = f"Loaded {len(tabs)} tabs"
        if combo_state == 'disabled':
            summary = "Only one tab available"
        self.tab_hint.config(text=summary)
        return self.selected_tab_gid

    def on_tab_selected(self, _event=None):
        selected_title = self.sheet_tab_var.get()
        for tab in self.sheet_tabs:
            if tab['title'] == selected_title:
                self.selected_tab_gid = tab['gid']
                self.gid_var.set(self.selected_tab_gid)
                self.working_csv_url = None
                break

    def report_environment(self):
        self.log(f"SheetDL v{VERSION} • https://github.com/{GITHUB_REPO}")
        ytdlp_version = getattr(yt_dlp, '__version__', 'unknown')
        ffmpeg_path = shutil.which('ffmpeg')
        if ffmpeg_path:
            self.log(f"Media tools • yt-dlp {ytdlp_version}, FFmpeg detected")
        else:
            self.log(f"Media tools • yt-dlp {ytdlp_version}, FFmpeg missing (downloads stay in source format)")
        
    def browse_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder_var.set(folder)
            
    def save_settings(self):
        """Save current settings"""
        self.config["sheet_url"] = self.sheet_url_var.get()
        self.config["gid"] = self.gid_var.get()
        self.config["output_folder"] = self.output_folder_var.get()
        self.config["organize_by"] = self.organize_var.get()
        self.config["create_zip"] = self.create_zip_var.get()
        self.config["save_metadata"] = self.save_metadata_var.get()
        self.config["save_log"] = self.save_log_var.get()
        self.config["yt_format"] = self.yt_format_var.get()
        self.config["sc_format"] = self.sc_format_var.get()
        self.config["column_mapping"] = {
            "artist": self.artist_col_var.get(),
            "title": self.title_col_var.get(),
            "url": self.url_col_var.get(),
            "genre": self.genre_col_var.get(),
            "cover": self.cover_col_var.get()
        }
        self.save_config()
        self.log("Settings saved successfully!")
        messagebox.showinfo("Success", "Settings saved!")
    
    def check_for_updates_gui(self):
        """Check for updates (GUI button callback)"""
        self.log("Checking for updates...")
        check_for_updates(silent=False)
    
    def add_to_queue(self):
        """Add current sheet URL to the download queue"""
        url = self.sheet_url_var.get()
        gid = self.gid_var.get() or "0"
        
        if not url:
            messagebox.showerror("Error", "Please enter a Google Sheets URL to add to queue")
            return
        
        # Get sheet name for display
        try:
            sheet_id = self.extract_sheet_id(url)
            name = self.get_sheet_title(url, sheet_id) or f"Sheet {len(self.download_queue) + 1}"
        except:
            name = f"Sheet {len(self.download_queue) + 1}"
        
        # Check for duplicates
        for item in self.download_queue:
            if item['url'] == url and item['gid'] == gid:
                messagebox.showinfo("Already in Queue", "This sheet is already in the queue.")
                return
        
        self.download_queue.append({'url': url, 'gid': gid, 'name': name})
        self.update_queue_display()
        self.log(f"Added to queue: {name}")
        
        # Clear the URL field so user can add another
        self.sheet_url_var.set("")
        self.gid_var.set("0")
    
    def update_queue_display(self):
        """Update the queue listbox display"""
        self.queue_listbox.delete(0, tk.END)
        
        if self.download_queue:
            self.queue_frame.grid()  # Show queue frame
            self.queue_label.config(text=f"📋 Queue ({len(self.download_queue)})")
            for i, item in enumerate(self.download_queue):
                self.queue_listbox.insert(tk.END, f"  {i+1}. {item['name']}")
        else:
            self.queue_frame.grid_remove()  # Hide queue frame
    
    def remove_queue_item(self, event=None):
        """Remove selected item from queue (double-click)"""
        selection = self.queue_listbox.curselection()
        if selection:
            idx = selection[0]
            removed = self.download_queue.pop(idx)
            self.log(f"Removed from queue: {removed['name']}")
            self.update_queue_display()
    
    def clear_queue(self):
        """Clear all items from the queue"""
        if self.download_queue:
            if messagebox.askyesno("Clear Queue", f"Remove all {len(self.download_queue)} items from the queue?"):
                self.download_queue.clear()
                self.update_queue_display()
                self.log("Queue cleared")
    
    def toggle_pause(self):
        """Toggle pause/resume state"""
        if self.is_paused:
            self.is_paused = False
            self.pause_btn.update_text("Pause")
            self.log("▶ Resumed downloading...")
        else:
            self.is_paused = True
            self.pause_btn.update_text("Resume")
            self.log("⏸ Paused - click Resume to continue")
    
    def process_next_queue_item(self):
        """Process the next item in the queue"""
        if self.download_queue and not self.is_downloading:
            next_item = self.download_queue.pop(0)
            self.update_queue_display()
            
            # Set the URL and GID
            self.sheet_url_var.set(next_item['url'])
            self.gid_var.set(next_item['gid'])
            
            self.log(f"Starting next in queue: {next_item['name']}")
            
            # Start download
            self.start_download()
        
    def test_connection(self):
        """Test Google Sheets connection"""
        self.log("Testing connection to Google Sheets...")
        try:
            sheet_url = self.sheet_url_var.get()
            if not sheet_url:
                messagebox.showerror("Error", "Please enter a Google Sheets URL")
                return
                
            # Try to access the sheet (public read-only)
            sheet_id = self.extract_sheet_id(sheet_url)
            if not sheet_id:
                messagebox.showerror("Error", "Invalid Google Sheets URL")
                return
                
            # Extract gid (sheet tab ID) if present or use the configured value
            preferred_gid = self.gid_var.get() or "0"
            gid_match = re.search(r'gid=([0-9]+)', sheet_url)
            if gid_match:
                preferred_gid = gid_match.group(1)

            tabs = self.fetch_sheet_tabs(sheet_id)
            gid = self.update_sheet_tab_options(tabs, preferred_gid)
            
            # Try multiple methods to access the sheet
            methods = [
                # Method 1: Direct pub URL (works for "Publish to web" sheets)
                f"https://docs.google.com/spreadsheets/d/e/{sheet_id}/pub?output=csv&gid={gid}",
                # Method 2: Standard export URL
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}",
                # Method 3: Direct gviz query
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}",
            ]
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
            
            success = False
            for idx, csv_url in enumerate(methods):
                self.log(f"Trying method {idx+1}...")
                try:
                    response = requests.get(csv_url, timeout=10, headers=headers, allow_redirects=True)
                    
                    if response.status_code == 200 and len(response.text) > 50:
                        lines = response.text.strip().split('\n')
                        if len(lines) > 0:
                            self.log(f"✓ Connection successful! Found {len(lines)-1} rows")
                            self.log(f"Headers: {lines[0]}")
                            # Save the working method
                            self.working_csv_url = csv_url
                            messagebox.showinfo("Success", f"Connected successfully!\nFound {len(lines)-1} rows\n\nHeaders: {lines[0][:100]}...")
                            success = True
                            break
                except:
                    continue
            
            if not success:
                self.log("✗ All connection methods failed")
                messagebox.showwarning("Connection Issue", 
                    "Cannot auto-access the sheet.\n\n" +
                    "Please make it PUBLIC:\n" +
                    "1. Open your sheet\n" +
                    "2. File → Share → Publish to web\n" +
                    "3. Click 'Publish'\n\n" +
                    "OR share it as:\n" +
                    "1. Click Share button\n" +
                    "2. Change to 'Anyone with the link'\n" +
                    "3. Set permission to 'Viewer'\n\n" +
                    "Then try again!")
                
        except Exception as e:
            self.log(f"✗ Connection failed: {str(e)}")
            messagebox.showerror("Error", f"Connection failed: {str(e)}")
            
    def extract_sheet_id(self, url):
        """Extract sheet ID from Google Sheets URL"""
        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
        return match.group(1) if match else None
    
    def extract_embedded_hyperlinks(self, sheet_id, gid):
        """Extract hyperlinks embedded in cells from Google Sheets view page.
        
        This is needed for sheets where URLs are in HYPERLINK() formulas
        but the display text is just the format (e.g., 'MP3', 'WAV').
        The CSV export only returns the display text, not the underlying URL.
        
        Returns a list of URLs found in the sheet.
        """
        try:
            view_url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/view?gid={gid}'
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
            response = requests.get(view_url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                return []
            
            # Extract all URLs from the page source
            # Common patterns for file hosting sites
            url_patterns = [
                r'https?://krakenfiles\.com/view/[a-zA-Z0-9]+/file\.html',
                r'https?://(?:www\.)?pillows\.su/[^\s"\'\\<>]+',
                r'https?://plwcse\.top/[^\s"\'\\<>]+',
                r'https?://pixeldrain\.com/[^\s"\'\\<>]+',
                r'https?://(?:www\.)?mega\.nz/[^\s"\'\\<>]+',
                r'https?://drive\.google\.com/[^\s"\'\\<>]+',
                r'https?://(?:www\.)?fileditch[^\s"\'\\<>]+',
                r'https?://music\.froste\.lol/[^\s"\'\\<>]+',
                r'https?://(?:www\.)?bumpworthy\.com/[^\s"\'\\<>]+',
                r'https?://(?:www\.)?soundcloud\.com/[^\s"\'\\<>]+',
                r'https?://(?:www\.)?youtube\.com/[^\s"\'\\<>]+',
                r'https?://youtu\.be/[^\s"\'\\<>]+',
                r'https?://(?:i\.)?imgur\.com/[^\s"\'\\<>]+',
                r'https?://(?:i\.)?imgur\.gg/[^\s"\'\\<>]+',
                r'https?://gofile\.io/d/[a-zA-Z0-9]+',
                r'https?://(?:www\.)?mediafire\.com/file/[^\s"\'\\<>]+',
                r'https?://[^\s"\'\\<>]*\.?s3[^\s"\'\\<>]*\.amazonaws\.com/[^\s"\'\\<>]+',
            ]
            
            all_urls = []
            for pattern in url_patterns:
                matches = re.findall(pattern, response.text, re.IGNORECASE)
                for url in matches:
                    # Clean up escaped characters
                    clean_url = url.replace('\\u002F', '/').replace('\\/', '/').rstrip('\\').rstrip(',').rstrip('"')
                    if clean_url not in all_urls:
                        all_urls.append(clean_url)
            
            return all_urls
            
        except Exception as e:
            return []
        
    def start_download(self):
        """Start the download process"""
        if self.is_downloading:
            return
            
        # Validate settings
        if not self.sheet_url_var.get():
            messagebox.showerror("Error", "Please enter a Google Sheets URL")
            return
            
        if not self.output_folder_var.get():
            messagebox.showerror("Error", "Please select an output folder")
            return
            
        # Create output folder
        os.makedirs(self.output_folder_var.get(), exist_ok=True)
        
        # Update UI
        self.is_downloading = True
        self.is_paused = False
        self.download_btn.set_state('disabled')
        self.pause_btn.set_state('normal')
        self.pause_btn.update_text("Pause")
        self.stop_btn.set_state('normal')
        self.queue_btn.set_state('normal')  # Enable adding to queue while downloading
        self.progress_var.set(0)
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # Start download in separate thread
        self.download_thread = threading.Thread(target=self.download_process, daemon=True)
        self.download_thread.start()
        
    def stop_download(self):
        """Stop the download process"""
        self.is_downloading = False
        self.is_paused = False
        self.log("Stopping download...")
        
    def download_process(self):
        """Main download process"""
        try:
            self.log("Starting download process...")
            
            # Get sheet data
            sheet_url = self.sheet_url_var.get()
            sheet_id = self.extract_sheet_id(sheet_url)
            
            # Extract gid (sheet tab ID) if present
            gid = self.gid_var.get() or "0"
            gid_match = re.search(r'gid=([0-9]+)', sheet_url)
            if gid_match:
                gid = gid_match.group(1)
                self.gid_var.set(gid)

            # Discover sheet title for folder naming
            self.current_sheet_name = self.get_sheet_title(sheet_url, sheet_id)
            
            # Use the working CSV URL if we found one, otherwise try all methods
            csv_urls = []
            if self.working_csv_url:
                csv_urls = [self.working_csv_url]
            else:
                csv_urls = [
                    f"https://docs.google.com/spreadsheets/d/e/{sheet_id}/pub?output=csv&gid={gid}",
                    f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}",
                    f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}",
                ]
                
            self.log("Fetching sheet data...")
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
            
            response = None
            for csv_url in csv_urls:
                try:
                    response = requests.get(csv_url, timeout=30, headers=headers, allow_redirects=True)
                    if response.status_code == 200 and len(response.text) > 50:
                        self.log("✓ Successfully fetched sheet data")
                        break
                except:
                    continue
            
            if not response or response.status_code != 200 or len(response.text) == 0:
                self.log(f"✗ Failed to fetch sheet data")
                messagebox.showerror("Error", "Cannot access sheet. Please use 'Test Connection' first to verify access.")
                return
                
            # Parse CSV
            import csv
            from io import StringIO
            
            all_rows = list(csv.reader(StringIO(response.text)))
            if not all_rows:
                self.log("✗ Sheet returned no rows")
                messagebox.showerror("Error", "Sheet appears to be empty.")
                return
            
            # The first row often contains stacked header text with data
            # Example: "Era 47 Full 0 Tagged..." where "Era" is the header
            # We need to extract just the header name from each cell
            def extract_header_name(cell_text):
                """Extract just the header name from a stacked cell"""
                if not cell_text:
                    return ""
                first_line = cell_text.split('\n')[0].strip()
                # Known header patterns to extract
                header_patterns = [
                    (r'^(Era)\b', 'Era'),
                    (r'^(Name)\b', 'Name'),
                    (r'^(Notes?)\b', 'Notes'),
                    (r'^(Track Length)\b', 'Track Length'),
                    (r'^(File Date)\b', 'File Date'),
                    (r'^(Leak Date)\b', 'Leak Date'),
                    (r'^(Type)\b', 'Type'),
                    (r'^(Available)\b', 'Available'),
                    (r'^(Quality)\b', 'Quality'),
                    (r'^(Link\(?s?\)?)\b', 'Link(s)'),
                    (r'^(Artist)\b', 'Artist'),
                    (r'^(Title)\b', 'Title'),
                    (r'^(Album)\b', 'Album'),
                    (r'^(Genre)\b', 'Genre'),
                    (r'^(Cover)\b', 'Cover'),
                    (r'^(URL)\b', 'URL'),
                    (r'^(Download)\b', 'Download'),
                    (r'^(Project)\b', 'Project'),
                ]
                for pattern, name in header_patterns:
                    if re.match(pattern, first_line, re.IGNORECASE):
                        return name
                # Fallback: return first word if it looks like a header (short, no URLs)
                if len(first_line) < 50 and 'http' not in first_line.lower():
                    return first_line.split()[0] if first_line.split() else ""
                return ""

            raw_header_row = all_rows[0]
            clean_headers = [extract_header_name(cell) for cell in raw_header_row]
            
            # Make sure we have at least column placeholders
            for i, h in enumerate(clean_headers):
                if not h:
                    clean_headers[i] = f"Column {i+1}"
            
            self.log(f"Detected columns: {[h for h in clean_headers if not h.startswith('Column')]}")
            
            # Build rows as dictionaries using the clean headers
            rows = []
            for row_data in all_rows[1:]:
                row_dict = {}
                for i, cell in enumerate(row_data):
                    if i < len(clean_headers):
                        header_name = clean_headers[i]
                        # For data cells, take only the first line
                        first_line = cell.split('\n')[0].strip() if cell else ""
                        row_dict[header_name] = first_line
                if any((v or "").strip() for v in row_dict.values()):
                    rows.append(row_dict)
            
            total_rows = len(rows)
            self.log(f"Found {total_rows} usable rows")
            
            if total_rows == 0:
                messagebox.showwarning("No Data", "No rows with data were found after the header.")
                return
            
            headers_list = [h for h in clean_headers if h]
            
            # For this sheet format:
            # - "Name" column = Song title (with emoji like "⭐️ (Nice Dream) [V1]")
            # - "Era" column = Album/Project folder (like "The Bends", "OK Computer")
            # - "Link(s)" column = Download URLs
            
            def find_best_column(preferred_names):
                """Find column matching any of the preferred names (case-insensitive)"""
                for pref in preferred_names:
                    for col in headers_list:
                        if col.lower() == pref.lower():
                            return col
                return None
            
            # Detect key columns
            title_col = find_best_column(['Name', 'Title', 'Track', 'Song'])
            album_col = find_best_column(['Era', 'Album', 'Project', 'Release'])
            url_col = find_best_column(['Link(s)', 'Links', 'Link', 'URL', 'Download'])
            artist_col = find_best_column(['Artist', 'Credited', 'Singer'])
            genre_col = find_best_column(['Genre', 'Category', 'Type'])
            cover_col = find_best_column(['Cover', 'Artwork', 'Image'])
            notes_col = find_best_column(['Notes', 'Note', 'Description', 'Info'])
            format_col = find_best_column(['Format', 'Quality', 'Media Type', 'Output'])
            type_col = find_best_column(['Type', 'Version', 'Status'])
            file_date_col = find_best_column(['File Date', 'Date', 'Recording Date'])
            leak_date_col = find_best_column(['Leak Date', 'Release Date', 'Leaked'])
            
            self.log(f"Column mapping: Title='{title_col}', Album='{album_col}', URL='{url_col}'")
            
            if not title_col:
                self.log("⚠ Could not find song name column (Name/Title)")
                title_col = headers_list[1] if len(headers_list) > 1 else headers_list[0]
                self.log(f"  Using column '{title_col}' as fallback")
            
            if not url_col:
                # Check which column actually contains URLs
                for col in headers_list:
                    for row in rows[:20]:
                        val = row.get(col, '')
                        if val and ('http' in val.lower() or 'pillows' in val.lower()):
                            url_col = col
                            self.log(f"Found URLs in column: '{col}'")
                            break
                    if url_col:
                        break
            
            # Check if URL column has actual URLs or just format text like "MP3", "WAV"
            embedded_hyperlinks = []
            urls_found_in_csv = False
            use_embedded_mode = False
            
            if url_col:
                for row in rows[:20]:
                    val = row.get(url_col, '')
                    if val and 'http' in val.lower():
                        urls_found_in_csv = True
                        break
            
            if not url_col or not urls_found_in_csv:
                # Try to extract embedded hyperlinks from the sheet view
                self.log("URL column contains format text (like 'MP3'), not actual URLs.")
                self.log("Extracting embedded hyperlinks from sheet...")
                embedded_hyperlinks = self.extract_embedded_hyperlinks(sheet_id, gid)
                
                if embedded_hyperlinks:
                    self.log(f"✓ Found {len(embedded_hyperlinks)} embedded download links")
                    use_embedded_mode = True
                else:
                    self.log("✗ Could not find any download URLs in sheet!")
                    messagebox.showerror("Error", "Could not find a column with download URLs.\n\nThis sheet may use hyperlinks embedded in cells (like clickable 'MP3' text).\nTry opening the sheet in your browser and copying the actual download URLs.")
                    return
            
            # Download each track
            success_count = 0
            fail_count = 0
            failed_downloads = []  # Track failed downloads for log file
            
            # If using embedded mode, download all embedded URLs with their original filenames
            if use_embedded_mode:
                self.log(f"\n=== Embedded Hyperlinks Mode ===")
                self.log(f"Downloading {len(embedded_hyperlinks)} files using original filenames from sources...")
                
                # Use sheet name as folder
                sheet_folder = self.sanitize_filename(self.current_sheet_name or "Sheet") or "Sheet"
                output_folder = Path(self.output_folder_var.get()) / sheet_folder
                os.makedirs(output_folder, exist_ok=True)
                
                for idx, url in enumerate(embedded_hyperlinks):
                    if not self.is_downloading:
                        self.log("Download stopped by user")
                        break
                    
                    self.progress_var.set(int((idx / len(embedded_hyperlinks)) * 100))
                    self.log(f"\n[{idx+1}/{len(embedded_hyperlinks)}] {url}")
                    
                    # Use a placeholder title - the download function will use the original filename
                    success = self.download_file(url, output_folder, self.current_sheet_name or "Unknown", f"Track_{idx+1}")
                    
                    if not success:
                        self.log("    ⟳ Retrying download...")
                        success = self.download_file(url, output_folder, self.current_sheet_name or "Unknown", f"Track_{idx+1}")
                    
                    if success:
                        success_count += 1
                    else:
                        fail_count += 1
                        failed_downloads.append({'title': url, 'url': url, 'error': 'Download failed'})
                
                self.progress_var.set(100)
                self.log(f"\n{'='*50}")
                self.log(f"Download complete!")
                self.log(f"  Successful: {success_count}")
                self.log(f"  Failed: {fail_count}")
                
                self.is_downloading = False
                self.download_btn.set_state('normal')
                self.stop_btn.set_state('disabled')
                return
            
            # Normal mode - process rows with URL column
            embedded_url_index = 0  # Track which embedded URL we're on (fallback)
            
            for idx, row in enumerate(rows):
                if not self.is_downloading:
                    self.log("Download stopped by user")
                    break
                
                # Handle pause state
                while self.is_paused and self.is_downloading:
                    import time
                    time.sleep(0.5)
                    
                if not self.is_downloading:
                    self.log("Download stopped by user")
                    break
                    
                try:
                    # Get values from the row using detected columns
                    title_value = row.get(title_col, "") if title_col else ""
                    album_value = row.get(album_col, "") if album_col else ""
                    artist_value = row.get(artist_col, "") if artist_col else ""
                    genre_value = row.get(genre_col, "") if genre_col else ""
                    cover_value = row.get(cover_col, "") if cover_col else ""
                    url_cell = (row.get(url_col, "") or "").strip()

                    # Clean the values - title is the song name (first line only)
                    title = self.clean_title(title_value)
                    # Capture extra title info (performer, alternate names, etc.)
                    title_extra_info = self.get_full_title_info(title_value)
                    album = self.clean_multiline_value(album_value) if album_value else ""
                    artist = self.clean_artist(artist_value) if artist_value else ""
                    genre = self.clean_multiline_value(genre_value) if genre_value else ""
                    cover_url = self.extract_cover_url(cover_value) if cover_value else ""
                    
                    # If artist is empty/meaningless, default to sheet name
                    if not self.is_meaningful_text(artist):
                        artist = self.current_sheet_name or "Unknown Artist"
                    
                    # If title is empty/meaningless, use a track number
                    if not self.is_meaningful_text(title):
                        title = f"Track {idx+1}"

                    # Extract URLs from the cell
                    urls_in_cell = self.extract_urls_from_cell(url_cell)
                    
                    # If no URLs in cell but we have embedded hyperlinks, try to use one
                    if not urls_in_cell and embedded_hyperlinks:
                        # Check if url_cell contains format indicators (MP3, WAV, FLAC, etc.)
                        format_indicators = ['mp3', 'wav', 'flac', 'm4a', 'aac', 'ogg', 'wma', 'aiff', 'alac']
                        cell_lower = url_cell.lower().strip()
                        if cell_lower in format_indicators or any(cell_lower.startswith(f) for f in format_indicators):
                            # This row likely has an embedded hyperlink
                            if embedded_url_index < len(embedded_hyperlinks):
                                urls_in_cell = [embedded_hyperlinks[embedded_url_index]]
                                embedded_url_index += 1
                    
                    if not urls_in_cell:
                        continue  # Skip silently if no URL
                    
                    self.log(f"\n[{idx+1}/{total_rows}] {title}")
                    if album:
                        self.log(f"  Album/Era: {album}")
                    if len(urls_in_cell) > 1:
                        self.log(f"  Found {len(urls_in_cell)} URLs")
                    
                    # Determine output folder - use album/era as the subfolder
                    sheet_folder = self.sanitize_filename(self.current_sheet_name or "Sheet") or "Sheet"
                    base_path = Path(self.output_folder_var.get()) / sheet_folder
                    
                    # Add album/era subfolder if available
                    if album:
                        album_folder = self.sanitize_filename(album)
                        if album_folder:
                            row_folder = base_path / album_folder
                        else:
                            row_folder = base_path
                    else:
                        row_folder = base_path
                    
                    os.makedirs(row_folder, exist_ok=True)
                    row_has_success = False
                    
                    for link_idx, url in enumerate(urls_in_cell, start=1):
                        display_url = url if len(url) <= 100 else f"{url[:100]}..."
                        prefix = f"  URL {link_idx}: " if len(urls_in_cell) > 1 else "  URL: "
                        self.log(f"{prefix}{display_url}")
                        
                        # Detect URL type
                        lowered = url.lower()
                        if 'pillows.su' in lowered or 'plwcse.top' in lowered or 'pillowcase.zip' in lowered or 'pillowcase.su' in lowered:
                            alt_domain = ''
                            if 'plwcse.top' in lowered:
                                alt_domain = ' (from plwcse.top)'
                            elif 'pillowcase.zip' in lowered:
                                alt_domain = ' (from pillowcase.zip)'
                            elif 'pillowcase.su' in lowered:
                                alt_domain = ' (from pillowcase.su)'
                            self.log("    Type: pillows.su" + alt_domain)
                        elif 'krakenfiles.com' in lowered:
                            self.log("    Type: KrakenFiles")
                        elif 'music.froste.lol' in lowered or 'froste.lol/song' in lowered:
                            self.log("    Type: Froste.lol")
                        elif 'pixeldrain.com' in lowered:
                            self.log("    Type: Pixeldrain")
                        elif 'fileditch' in lowered or 'fileditchfiles' in lowered:
                            self.log("    Type: FileDitch")
                        elif 'bumpworthy.com' in lowered:
                            self.log("    Type: BumpWorthy")
                        elif 'drive.google.com' in lowered or 'docs.google.com' in lowered:
                            self.log("    Type: Google Drive")
                        elif 'mega.nz' in lowered or 'mega.co.nz' in lowered:
                            self.log("    Type: MEGA.nz")
                        elif 'imgur.com' in lowered or 'i.imgur.com' in lowered:
                            self.log("    Type: Imgur")
                        elif 'imgur.gg' in lowered or 'i.imgur.gg' in lowered:
                            self.log("    Type: imgur.gg")
                        elif 'dump.li' in lowered:
                            self.log("    Type: Dump.li")
                        elif 'catbox.moe' in lowered or 'files.catbox.moe' in lowered:
                            self.log("    Type: Catbox.moe")
                        elif 'ibb.co' in lowered or 'i.ibb.co' in lowered:
                            self.log("    Type: ibb.co")
                        elif 'gofile.io' in lowered:
                            self.log("    Type: Gofile.io")
                        elif 'mediafire.com' in lowered:
                            self.log("    Type: MediaFire")
                        elif 's3.amazonaws.com' in lowered or '.s3.' in lowered and 'amazonaws.com' in lowered:
                            self.log("    Type: AWS S3")
                        elif any(x in lowered for x in ['youtube.com', 'youtu.be']):
                            self.log("    Type: YouTube")
                        elif 'soundcloud.com' in lowered or 'on.soundcloud.com' in lowered:
                            self.log("    Type: SoundCloud")
                        else:
                            self.log("    Type: Direct download")
                        
                        link_title = title if len(urls_in_cell) == 1 else f"{title} (Link {link_idx})"
                        
                        # Try download with retry on failure
                        success = self.download_file(url, row_folder, artist, link_title)
                        
                        if not success:
                            self.log("    ⟳ Retrying download...")
                            success = self.download_file(url, row_folder, artist, link_title)
                        
                        if success:
                            success_count += 1
                            row_has_success = True
                            self.log("    ✓ SUCCESS")
                        else:
                            fail_count += 1
                            failed_downloads.append({
                                'title': link_title,
                                'artist': artist,
                                'url': url,
                                'row': idx + 1
                            })
                            self.log("    ✗ FAILED (after retry)")
                    
                    self.log("")

                    if not row_has_success:
                        continue

                    # Download cover art if available
                    cover_saved_name = None
                    if cover_url:
                        cover_path = self.download_album_cover(cover_url, row_folder, album or title)
                        if cover_path:
                            cover_saved_name = cover_path.name

                    # Create metadata summary file if enabled
                    if self.save_metadata_var.get():
                        metadata_filename = f"{self.build_safe_title(title)}.txt"
                        metadata_path = self.resolve_duplicate_path(row_folder / metadata_filename)

                        def column_value(column_name):
                            return self.clean_multiline_value(row.get(column_name, "")) if column_name else ""

                        notes_value = column_value(notes_col)
                        file_date_value = column_value(file_date_col)
                        leak_date_value = column_value(leak_date_col)
                        type_value = column_value(type_col)
                        format_value = column_value(format_col)

                        if not file_date_value:
                            file_date_value = self.find_first_date_in_row(row, exclude_columns=[notes_col, title_col])

                        metadata_lines = [
                            f"Title: {title}",
                        ]
                        
                        # Add extra title info if present (performer, alternate names, etc.)
                        if title_extra_info:
                            metadata_lines.append(f"Additional Info:\n{title_extra_info}")
                        
                        metadata_lines.extend([
                            f"Artist: {artist}",
                            f"Album/Project: {album or 'N/A'}",
                            f"Genre/Category: {genre or 'N/A'}",
                            f"Notes: {notes_value or 'N/A'}",
                            f"File Date: {file_date_value or 'N/A'}",
                            f"Leak/Release Date: {leak_date_value or 'N/A'}",
                            f"Type: {type_value or 'N/A'}",
                            f"Format: {format_value or 'N/A'}",
                            f"Cover Source: {cover_url or 'N/A'}",
                            f"Cover Saved: {cover_saved_name or 'N/A'}",
                            "Download Links:"
                        ])

                        if urls_in_cell:
                            metadata_lines.extend([f"  - {link}" for link in urls_in_cell])
                        else:
                            metadata_lines.append("  - None")

                        metadata_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

                        with open(metadata_path, 'w', encoding='utf-8') as meta_file:
                            meta_file.write('\n\n'.join(metadata_lines))
                        
                except Exception as e:
                    self.log(f"✗ Error processing row {idx+1}: {str(e)}")
                    fail_count += 1
                    
                # Update progress
                progress = ((idx + 1) / total_rows) * 100
                self.progress_var.set(progress)
                
            self.log(f"\n{'='*50}")
            self.log(f"Download complete!")
            self.log(f"Success: {success_count} | Failed: {fail_count}")
            
            # Log failed downloads summary
            if failed_downloads:
                self.log(f"\n{'='*50}")
                self.log("FAILED DOWNLOADS:")
                for item in failed_downloads:
                    self.log(f"  Row {item['row']}: {item['artist']} - {item['title']}")
                    self.log(f"    URL: {item['url']}")
            
            # Save log file if enabled
            if self.save_log_var.get():
                self._save_download_log(success_count, fail_count, failed_downloads)
            
            # Create ZIP if requested
            if self.create_zip_var.get() and success_count > 0:
                self.log("\nCreating ZIP archive...")
                self.create_zip_archive()
                
            messagebox.showinfo("Complete", f"Download finished!\nSuccess: {success_count}\nFailed: {fail_count}")
            
        except Exception as e:
            self.log(f"✗ Fatal error: {str(e)}")
            messagebox.showerror("Error", f"Download failed: {str(e)}")
            
        finally:
            self.is_downloading = False
            self.is_paused = False
            self.download_btn.set_state('normal')
            self.pause_btn.set_state('disabled')
            self.pause_btn.update_text("Pause")
            self.stop_btn.set_state('disabled')
            
            # Check if there are more items in the queue
            if self.download_queue:
                self.log(f"📋 {len(self.download_queue)} more sheet(s) in queue...")
                self.root.after(1000, self.process_next_queue_item)
            else:
                self.queue_btn.set_state('disabled')
    
    def _save_download_log(self, success_count, fail_count, failed_downloads):
        """Save the download log to a text file"""
        try:
            sheet_folder = self.sanitize_filename(self.current_sheet_name or "Sheet") or "Sheet"
            base_path = Path(self.output_folder_var.get()) / sheet_folder
            os.makedirs(base_path, exist_ok=True)
            
            log_filename = f"download_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            log_path = base_path / log_filename
            
            # Get the full log content from the text widget
            self.log_text.config(state='normal')
            log_content = self.log_text.get(1.0, tk.END)
            self.log_text.config(state='disabled')
            
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write(f"SheetDL Download Log\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*60}\n\n")
                f.write(f"SUMMARY\n")
                f.write(f"Success: {success_count}\n")
                f.write(f"Failed: {fail_count}\n\n")
                
                if failed_downloads:
                    f.write(f"{'='*60}\n")
                    f.write("FAILED DOWNLOADS:\n\n")
                    for item in failed_downloads:
                        f.write(f"Row {item['row']}: {item['artist']} - {item['title']}\n")
                        f.write(f"  URL: {item['url']}\n\n")
                
                f.write(f"{'='*60}\n")
                f.write("FULL LOG:\n\n")
                f.write(log_content)
            
            self.log(f"\n✓ Download log saved: {log_path.name}")
        except Exception as e:
            self.log(f"✗ Failed to save log file: {str(e)}")
            
    def get_output_path(self, artist, genre):
        """Determine output path based on organization settings"""
        sheet_folder = self.sanitize_filename(self.current_sheet_name or "Sheet") or "Sheet"
        base_path = Path(self.output_folder_var.get()) / sheet_folder
        organize_by = self.organize_var.get()
        
        if organize_by == "artist":
            return base_path / self.sanitize_filename(artist)
        elif organize_by == "genre" and genre:
            return base_path / self.sanitize_filename(genre)
        elif organize_by == "artist_genre" and genre:
            return base_path / self.sanitize_filename(genre) / self.sanitize_filename(artist)
        else:
            return base_path
            
    def sanitize_filename(self, filename):
        """Remove or replace invalid characters from filename"""
        if not filename:
            return ""
        # Replace question marks with a similar-looking character (to preserve "??? [V1]" style names)
        result = filename.replace('?', '¿')
        # Remove other invalid characters
        result = re.sub(r'[<>:"/\\|*]', '', result)
        return result.strip()

    def clean_multiline_value(self, value, pick_last=False):
        """Normalize multi-line cell values, optionally keeping last non-empty line"""
        if value is None:
            return ""
        text = str(value).replace('\r', '\n')
        parts = [part.strip() for part in text.split('\n') if part.strip()]
        if not parts:
            return ""
        choice = parts[-1] if pick_last else parts[0]
        return re.sub(r'\s+', ' ', choice).strip()

    def is_meaningful_text(self, text):
        stripped = str(text or "").strip()
        if not stripped:
            return False
        lowered = stripped.lower()
        return lowered not in {"unknown", "untitled", "n/a", "na", "tbd"}

    def clean_title(self, raw_title):
        cleaned = self.clean_multiline_value(raw_title, pick_last=False)
        return cleaned or "Untitled"

    def clean_artist(self, raw_artist):
        cleaned = self.clean_multiline_value(raw_artist, pick_last=False)
        return cleaned or "Unknown Artist"

    def get_full_title_info(self, raw_title):
        """Get all lines from a multi-line title cell (performer info, alternate names, etc.)"""
        if raw_title is None:
            return ""
        text = str(raw_title).replace('\r', '\n')
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if not lines:
            return ""
        # Skip the first line (that's the main title), return the rest
        extra_lines = lines[1:] if len(lines) > 1 else []
        return '\n'.join(extra_lines)

    def build_safe_title(self, title):
        return self.sanitize_filename(self.clean_title(title)) or "Track"

    def build_track_filename(self, title, link_suffix=None, extension=".mp3"):
        base_title = self.clean_title(title)
        if link_suffix:
            base_title = f"{base_title} ({link_suffix})"
        safe = self.sanitize_filename(base_title) or "Track"
        return f"{safe}{extension}"

    def find_first_date_in_row(self, row, exclude_columns=None):
        """Find first date-like value in row, returning only the date portion"""
        exclude_columns = exclude_columns or []
        date_pattern = re.compile(r'\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b')
        month_date_pattern = re.compile(r'\b((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})\b', re.IGNORECASE)
        year_pattern = re.compile(r'\b(\d{4})\b')
        
        for col_name, value in row.items():
            # Skip excluded columns (like Notes)
            if col_name in exclude_columns:
                continue
            text = str(value or '').strip()
            if not text or len(text) > 100:  # Skip long text fields
                continue
            # Try to match specific date formats
            match = date_pattern.search(text)
            if match:
                return match.group(1)
            match = month_date_pattern.search(text)
            if match:
                return match.group(1)
        
        # Fallback: look for just a year in short fields
        for col_name, value in row.items():
            if col_name in exclude_columns:
                continue
            text = str(value or '').strip()
            if not text or len(text) > 20:  # Only check short fields for year-only
                continue
            match = year_pattern.search(text)
            if match:
                return match.group(1)
        return ""

    def resolve_duplicate_path(self, filepath):
        filepath = Path(filepath)
        if not filepath.exists():
            return filepath
        base = filepath.stem
        suffix = filepath.suffix
        counter = 2
        while True:
            candidate = filepath.parent / f"{base} ({counter}){suffix}"
            if not candidate.exists():
                return candidate
            counter += 1

    def get_sheet_title(self, sheet_url, sheet_id):
        """Best-effort fetch of sheet title for folder naming"""
        default = "Sheet"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
        candidates = []

        if sheet_url:
            base = sheet_url.split('#')[0]
            candidates.append(base)
        if sheet_id:
            candidates.extend([
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/view",
                f"https://docs.google.com/spreadsheets/d/{sheet_id}/pubhtml"
            ])

        for target in candidates:
            try:
                resp = requests.get(target, headers=headers, timeout=10)
                if resp.status_code != 200:
                    continue
                title = self.extract_sheet_title(resp.text)
                if title:
                    title = title.replace(' - Google Sheets', '').strip()
                    title = re.sub(r'(?i)tracker', '', title).strip(' -_')
                    if title:
                        return title
            except Exception:
                continue

        return default

    def extract_sheet_title(self, html_text):
        try:
            soup = BeautifulSoup(html_text, 'html.parser')
            if soup.title and soup.title.string:
                title = soup.title.string.strip()
                if title and 'google accounts' not in title.lower():
                    return title
            og_title = soup.find('meta', attrs={'property': 'og:title'})
            if og_title and og_title.get('content'):
                return og_title['content'].strip()
        except Exception:
            pass
        return None

    def infer_extension(self, content_type, default=".mp3"):
        if not content_type:
            return default
        content_type = content_type.lower()
        if 'flac' in content_type:
            return '.flac'
        if 'wav' in content_type or 'wave' in content_type:
            return '.wav'
        if 'ogg' in content_type:
            return '.ogg'
        if 'm4a' in content_type:
            return '.m4a'
        if 'aac' in content_type:
            return '.aac'
        if 'mp3' in content_type or 'mpeg' in content_type:
            return '.mp3'
        return default

    def extract_urls_from_cell(self, cell_value):
        """Extract all URLs from a spreadsheet cell"""
        if not cell_value:
            return []
        text = str(cell_value).replace('\n', ' ').replace('\r', ' ').strip()
        if not text:
            return []
        url_pattern = re.compile(r'(https?://[^\s,]+)', re.IGNORECASE)
        urls = url_pattern.findall(text)

        # Handle pillows.su links that might miss the scheme
        if not urls and 'pillows.su' in text.lower():
            alt_pattern = re.compile(r'((?:https?://)?pillows\.su/[^\s,]+)', re.IGNORECASE)
            urls = alt_pattern.findall(text)

        cleaned = []
        seen = set()
        for url in urls:
            trimmed = url.strip().rstrip(').,;\'"')
            if not trimmed.lower().startswith('http'):
                trimmed = f"https://{trimmed}"
            normalized = trimmed.rstrip('/')
            if normalized.lower() in seen:
                continue
            seen.add(normalized.lower())
            cleaned.append(trimmed)
        return cleaned

    def extract_cover_url(self, cell_value):
        urls = self.extract_urls_from_cell(cell_value)
        for url in urls:
            if self.is_image_url(url):
                return url
        return urls[0] if urls else ""

    def is_image_url(self, url):
        if not url:
            return False
        path = urlparse(url).path.lower()
        if any(path.endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.webp']):
            return True
        lowered = url.lower()
        return any(host in lowered for host in ['googleusercontent', 'imgur', 'ibb.co', 'postimg', 'cdn.discordapp.com'])

    def infer_image_extension(self, content_type, url=""):
        ext = Path(urlparse(url).path).suffix.lower()
        if ext in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
            return ext
        if content_type:
            content_type = content_type.lower()
            if 'png' in content_type:
                return '.png'
            if 'gif' in content_type:
                return '.gif'
            if 'webp' in content_type:
                return '.webp'
            if 'jpeg' in content_type or 'jpg' in content_type:
                return '.jpg'
        return '.jpg'

    def find_existing_cover(self, folder):
        folder = Path(folder)
        if not folder.exists():
            return None
        for candidate in folder.glob('cover.*'):
            if candidate.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
                return candidate
        for name in ('folder', 'front'):
            for candidate in folder.glob(f"{name}.*"):
                if candidate.suffix.lower() in ['.jpg', '.jpeg', '.png', '.gif', '.webp']:
                    return candidate
        return None

    def download_album_cover(self, url, target_folder, album_label):
        try:
            if not url:
                return None
            target_folder = Path(target_folder)
            os.makedirs(target_folder, exist_ok=True)
            folder_key = str(target_folder.resolve())
            if folder_key in self.cover_cache:
                existing_cached = self.find_existing_cover(target_folder)
                if existing_cached:
                    return existing_cached
            existing = self.find_existing_cover(target_folder)
            if existing:
                self.cover_cache.add(folder_key)
                return existing

            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            extension = self.infer_image_extension(response.headers.get('Content-Type'), url)
            filename = f"cover{extension}"
            cover_path = self.resolve_duplicate_path(target_folder / filename)
            with open(cover_path, 'wb') as image_file:
                for chunk in response.iter_content(chunk_size=4096):
                    if not chunk:
                        continue
                    image_file.write(chunk)
            self.cover_cache.add(folder_key)
            self.log(f"  ✓ Saved cover art for {album_label}: {cover_path.name}")
            return cover_path
        except Exception as e:
            self.log(f"  ✗ Cover download failed: {str(e)}")
            return None
        
    def download_file(self, url, output_path, artist, title):
        """Download a single file."""
        try:
            os.makedirs(output_path, exist_ok=True)
            
            # Convert legacy domain URLs to pillows.su
            if 'plwcse.top' in url:
                url = url.replace('plwcse.top', 'pillows.su')
                self.log(f"    → Converted plwcse.top URL to pillows.su")
            elif 'pillowcase.zip' in url:
                url = url.replace('pillowcase.zip', 'pillows.su')
                self.log(f"    → Converted pillowcase.zip URL to pillows.su")
            elif 'pillowcase.su' in url:
                url = url.replace('pillowcase.su', 'pillows.su')
                self.log(f"    → Converted pillowcase.su URL to pillows.su")
            
            # Check what type of URL this is
            if 'pillows.su' in url:
                return self.download_pillows(url, output_path, artist, title)
            elif 'krakenfiles.com' in url:
                return self.download_krakenfiles(url, output_path, artist, title)
            elif 'music.froste.lol' in url or 'froste.lol/song' in url:
                return self.download_froste(url, output_path, artist, title)
            elif 'pixeldrain.com' in url:
                return self.download_pixeldrain(url, output_path, artist, title)
            elif 'fileditch' in url or 'fileditchfiles' in url:
                return self.download_fileditch(url, output_path, artist, title)
            elif 'bumpworthy.com' in url:
                return self.download_bumpworthy(url, output_path, artist, title)
            elif 'drive.google.com' in url or 'docs.google.com' in url:
                return self.download_google_drive(url, output_path, artist, title)
            elif 'mega.nz' in url or 'mega.co.nz' in url:
                return self.download_mega(url, output_path, artist, title)
            elif 'imgur.com' in url or 'i.imgur.com' in url:
                return self.download_imgur(url, output_path, artist, title)
            elif 'imgur.gg' in url or 'i.imgur.gg' in url:
                return self.download_imgurgg(url, output_path, artist, title)
            elif 'ibb.co' in url or 'i.ibb.co' in url:
                return self.download_ibb(url, output_path, artist, title)
            elif 'gofile.io' in url:
                return self.download_gofile(url, output_path, artist, title)
            elif 'mediafire.com' in url:
                return self.download_mediafire(url, output_path, artist, title)
            elif 's3.amazonaws.com' in url or ('.s3.' in url and 'amazonaws.com' in url):
                return self.download_aws_s3(url, output_path, artist, title)
            elif any(x in url for x in ['youtube.com', 'youtu.be', 'soundcloud.com', 'on.soundcloud.com']):
                return self.download_youtube(url, output_path, artist, title)
            else:
                return self.download_direct(url, output_path, artist, title)
                
        except Exception as e:
            self.log(f"  Error: {str(e)}")
            return False
            
    def download_youtube(self, url, output_path, artist, title):
        """Download from YouTube or similar platforms using yt-dlp"""
        base_filename = self.build_safe_title(title)
        
        # Use SoundCloud format for SoundCloud URLs, YouTube format for everything else
        if 'soundcloud.com' in url.lower():
            yt_format = self.sc_format_var.get() if hasattr(self, 'sc_format_var') else 'audio_m4a'
        else:
            yt_format = self.yt_format_var.get() if hasattr(self, 'yt_format_var') else 'video_mp4'
        
        ydl_opts = {
            'outtmpl': str(output_path / f"{base_filename}.%(ext)s"),
            'noplaylist': True,
            'quiet': True,
            'no_warnings': True,
        }
        
        ffmpeg_path = shutil.which('ffmpeg')
        if ffmpeg_path:
            ydl_opts['ffmpeg_location'] = ffmpeg_path
        
        # Configure format based on user selection
        if yt_format == 'audio_m4a':
            # M4A audio - no conversion needed
            ydl_opts['format'] = 'bestaudio[ext=m4a]/bestaudio[ext=mp3]/bestaudio[ext=aac]/bestaudio'
            if not ffmpeg_path:
                self.log("    Downloading as m4a (no FFmpeg needed)")
        
        elif yt_format == 'audio_mp3':
            # MP3 audio - requires FFmpeg for conversion
            if ffmpeg_path:
                ydl_opts['format'] = 'bestaudio/best'
                ydl_opts['postprocessors'] = [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': '192',
                }]
            else:
                self.log("    FFmpeg not found - falling back to m4a")
                ydl_opts['format'] = 'bestaudio[ext=m4a]/bestaudio[ext=mp3]/bestaudio'
        
        elif yt_format == 'video_mp4':
            # MP4 video with audio
            ydl_opts['format'] = 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best'
            if ffmpeg_path:
                ydl_opts['merge_output_format'] = 'mp4'
        
        elif yt_format == 'video_best':
            # Best quality video regardless of format
            ydl_opts['format'] = 'bestvideo+bestaudio/best'
            if ffmpeg_path:
                ydl_opts['merge_output_format'] = 'mp4'
        
        else:
            # Default fallback to m4a
            ydl_opts['format'] = 'bestaudio[ext=m4a]/bestaudio'

        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
            return True
        except Exception as e:
            self.log(f"  yt-dlp error: {str(e)}")
            return False
    
    def download_pillows(self, url, output_path, artist, title):
        """Download from pillows.su by scraping the download link"""
        try:
            self.log(f"  → Accessing pillows.su page...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': 'https://pillows.su/',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
            }
            
            # Get the page
            response = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
            response.raise_for_status()
            
            # Parse HTML to find download link
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Look for the API download link
            # Pattern: https://api.pillows.su/api/download/{id}
            download_link = None
            
            # Method 1: Find direct API link in href
            self.log(f"  → Searching for download link...")
            for link in soup.find_all('a', href=True):
                href = link.get('href', '')
                if 'api.pillows.su/api/download' in href:
                    download_link = href
                    self.log(f"  ✓ Found API link in page")
                    break
            
            # Method 2: Extract file ID from URL and construct API link
            if not download_link:
                file_id_match = re.search(r'/f/([a-f0-9A-F]+)', url)
                if file_id_match:
                    file_id = file_id_match.group(1)
                    download_link = f"https://api.pillows.su/api/download/{file_id}"
                    self.log(f"  ✓ Constructed API link from file ID: {file_id}")
            
            if not download_link:
                self.log(f"  ✗ Could not find/construct download link")
                return False
            
            # Download the actual file from the API link
            download_headers = headers.copy()
            download_headers['Referer'] = url
            download_headers['Accept'] = 'application/octet-stream,application/json;q=0.9,*/*;q=0.8'
            response = requests.get(download_link, headers=download_headers, stream=True, timeout=60)
            response.raise_for_status()
            
            size_kb = None
            if response.headers.get('Content-Length'):
                try:
                    size_kb = int(response.headers['Content-Length']) // 1024
                except ValueError:
                    size_kb = None
            if size_kb:
                self.log(f"  → Downloading file (~{size_kb} KB)...")
            else:
                self.log("  → Downloading file (size unknown)...")
            
            # Build friendly filename with inferred extension
            extension = Path(urlparse(download_link).path).suffix
            if not extension:
                extension = self.infer_extension(response.headers.get('Content-Type'))
            filename = self.build_track_filename(title, extension=extension)
            
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Save the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ pillows.su error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False

    def download_krakenfiles(self, url, output_path, artist, title):
        """Download from krakenfiles.com - tries to get original file with CF bypass, falls back to m4a audio"""
        try:
            self.log(f"  → Accessing krakenfiles page...")
            
            import time
            import random
            
            # Try to use cloudscraper to bypass Cloudflare
            try:
                import cloudscraper
                session = cloudscraper.create_scraper(
                    browser={
                        'browser': 'chrome',
                        'platform': 'windows',
                        'desktop': True
                    }
                )
                self.log(f"  → Using cloudscraper for Cloudflare bypass")
            except ImportError:
                # Fallback to regular requests if cloudscraper not available
                session = requests.Session()
                chrome_versions = ['120.0.0.0', '121.0.0.0', '122.0.0.0', '123.0.0.0', '124.0.0.0', '125.0.0.0']
                chrome_ver = random.choice(chrome_versions)
                headers = {
                    'User-Agent': f'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{chrome_ver} Safari/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                    'Accept-Language': 'en-US,en;q=0.9',
                }
                session.headers.update(headers)
            
            # Extract file hash from URL
            hash_match = re.search(r'/(?:view|embed-audio)/([a-zA-Z0-9]+)', url)
            if not hash_match:
                self.log(f"  ✗ Could not extract file hash from URL")
                return False
            
            file_hash = hash_match.group(1)
            
            # Human-like behavior: small random delay before first request
            time.sleep(random.uniform(0.5, 1.5))
            
            # First visit the main page to establish cookies (like a real user)
            session.get('https://krakenfiles.com/', timeout=30)
            time.sleep(random.uniform(0.3, 0.8))
            
            # Now visit the file page
            page_response = session.get(url, timeout=30)
            
            # Small delay like reading the page
            time.sleep(random.uniform(0.5, 1.2))
            
            # Update referer for subsequent requests
            session.headers.update({'Referer': url})
            
            # Get file info from JSON API
            json_url = f"https://krakenfiles.com/json/{file_hash}"
            session.headers.update({
                'Accept': 'application/json, text/plain, */*',
                'Sec-Fetch-Dest': 'empty',
                'Sec-Fetch-Mode': 'cors',
                'Sec-Fetch-Site': 'same-origin',
            })
            
            time.sleep(random.uniform(0.2, 0.5))
            response = session.get(json_url, timeout=30)
            
            if response.status_code != 200:
                self.log(f"  ✗ Failed to get file info: HTTP {response.status_code}")
                return False
            
            info = response.json()
            original_title = info.get('title', '')
            server_url = info.get('serverUrl', '')
            upload_date_raw = info.get('uploadDate', '')
            file_type = info.get('type', '')
            file_size = info.get('size', 'unknown')
            
            self.log(f"  → Found: {original_title} ({file_size})")
            
            # If title looks like a placeholder (Track_N), use the original filename from KrakenFiles
            if title.startswith('Track_') or title.startswith('Track ') or not title:
                # Use original title (without extension) as the filename
                title_from_file = re.sub(r'\.[a-zA-Z0-9]+$', '', original_title)
                if title_from_file:
                    title = title_from_file
                    self.log(f"  → Using original filename: {title}")
            
            # Convert upload date from "14.11.2023 20:07" to "14-11-2023"
            upload_date = upload_date_raw.split()[0].replace('.', '-') if upload_date_raw else ''
            
            # Get original file extension from title
            original_ext_match = re.search(r'\.([a-zA-Z0-9]+)$', original_title)
            original_extension = f'.{original_ext_match.group(1)}' if original_ext_match else None
            
            # Try to get the original file first (with human-like behavior to bypass CF)
            original_downloaded = False
            cf_blocked = False
            if original_extension:
                try:
                    self.log(f"  → Attempting original file download...")
                    
                    # Simulate clicking the download button - small delay
                    time.sleep(random.uniform(0.8, 1.5))
                    
                    # Try the download endpoint
                    download_api_url = f"https://krakenfiles.com/download/{file_hash}"
                    
                    # Reset headers for download request
                    session.headers.update({
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
                        'Sec-Fetch-Dest': 'document',
                        'Sec-Fetch-Mode': 'navigate',
                        'Sec-Fetch-Site': 'same-origin',
                        'Sec-Fetch-User': '?1',
                    })
                    
                    # Make a few "warm up" requests like scrolling/mouse movement would trigger
                    # This sometimes helps with CF challenge detection
                    time.sleep(random.uniform(0.1, 0.3))
                    
                    dl_response = session.get(download_api_url, stream=True, timeout=60, allow_redirects=True)
                    
                    # Check if we got redirected to a CF challenge or captcha
                    content_type = dl_response.headers.get('Content-Type', '')
                    
                    if dl_response.status_code == 200 and 'text/html' not in content_type.lower():
                        # Success! We got the actual file
                        size_kb = None
                        if dl_response.headers.get('Content-Length'):
                            try:
                                size_kb = int(dl_response.headers['Content-Length']) // 1024
                            except ValueError:
                                pass
                        if size_kb:
                            self.log(f"  → Downloading original (~{size_kb} KB)...")
                        
                        filename = self.build_track_filename(title, extension=original_extension)
                        filepath = self.resolve_duplicate_path(output_path / filename)
                        
                        with open(filepath, 'wb') as f:
                            for chunk in dl_response.iter_content(chunk_size=8192):
                                if not self.is_downloading:
                                    return False
                                f.write(chunk)
                        
                        self.log(f"  ✓ Saved original as: {filepath.name}")
                        return True
                    else:
                        cf_blocked = True
                        self.log(f"  → Cloudflare CAPTCHA required - falling back to m4a stream")
                        
                except Exception as e:
                    self.log(f"  → Original download failed ({str(e)}), trying m4a...")
            
            # Fallback: Use the m4a stream (no CAPTCHA required)
            if file_type == 'music' and upload_date:
                download_url = f"{server_url}/uploads/{upload_date}/{file_hash}/music.m4a"
                extension = '.m4a'
                if original_extension and original_extension.lower() != '.m4a':
                    self.log(f"  → Using m4a stream (original requires CAPTCHA)")
            elif file_type == 'music':
                # Fallback without date folder
                download_url = f"{server_url}/uploads/{file_hash}.m4a"
                extension = '.m4a'
            else:
                # For non-music files, try to find the source in embed page
                embed_url = f"https://krakenfiles.com/embed-audio/{file_hash}"
                embed_resp = session.get(embed_url, timeout=30)
                
                m4a_match = re.search(r"m4a:\s*['\"]([^'\"]+)['\"]", embed_resp.text)
                if m4a_match:
                    download_url = m4a_match.group(1)
                    if download_url.startswith('//'):
                        download_url = 'https:' + download_url
                    extension = '.m4a'
                else:
                    # Fallback
                    extension = original_extension or '.mp3'
                    if upload_date:
                        download_url = f"{server_url}/uploads/{upload_date}/{file_hash}/file"
                    else:
                        download_url = f"{server_url}/uploads/{file_hash}/file"
            
            # Small delay before download
            time.sleep(random.uniform(0.3, 0.7))
            
            # Download the file
            response = session.get(download_url, stream=True, timeout=60)
            response.raise_for_status()
            
            size_kb = None
            if response.headers.get('Content-Length'):
                try:
                    size_kb = int(response.headers['Content-Length']) // 1024
                except ValueError:
                    pass
            if size_kb:
                self.log(f"  → Downloading (~{size_kb} KB)...")
            
            # Build filename
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Save the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ krakenfiles error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False

    def download_froste(self, url, output_path, artist, title):
        """Download from music.froste.lol by extracting the song ID and fetching the file"""
        try:
            self.log(f"  → Accessing froste.lol page...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': 'https://music.froste.lol/',
                'Accept': 'audio/*,application/octet-stream,*/*;q=0.8'
            }
            
            # Extract song ID from URL
            # Pattern: https://music.froste.lol/song/{songId}
            song_match = re.search(r'/song/([a-fA-F0-9]+)', url)
            if not song_match:
                self.log(f"  ✗ Could not extract song ID from froste.lol URL")
                return False
            
            song_id = song_match.group(1)
            self.log(f"  → Song ID: {song_id}")
            
            # The direct file endpoint is /song/{songId}/file
            download_url = f"https://music.froste.lol/song/{song_id}/file"
            
            self.log(f"  → Downloading audio file...")
            response = requests.get(download_url, headers=headers, stream=True, timeout=60)
            response.raise_for_status()
            
            # Get file size if available
            size_kb = None
            if response.headers.get('Content-Length'):
                size_bytes = int(response.headers.get('Content-Length'))
                size_kb = size_bytes / 1024
                if size_kb > 1024:
                    self.log(f"  → File size: {size_kb/1024:.1f} MB")
                else:
                    self.log(f"  → File size: {size_kb:.0f} KB")
            
            # Determine extension from Content-Type or Content-Disposition
            content_type = response.headers.get('Content-Type', '')
            content_disposition = response.headers.get('Content-Disposition', '')
            
            extension = '.mp3'  # Default
            if 'flac' in content_type.lower():
                extension = '.flac'
            elif 'm4a' in content_type.lower() or 'mp4' in content_type.lower():
                extension = '.m4a'
            elif 'wav' in content_type.lower():
                extension = '.wav'
            elif 'ogg' in content_type.lower():
                extension = '.ogg'
            
            # Try to get original filename from Content-Disposition
            if content_disposition:
                filename_match = re.search(r'filename[*]?=["\']?([^"\';\n]+)', content_disposition)
                if filename_match:
                    original_name = filename_match.group(1).strip()
                    ext_from_name = Path(original_name).suffix
                    if ext_from_name:
                        extension = ext_from_name
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Save the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ froste.lol error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False

    def download_pixeldrain(self, url, output_path, artist, title):
        """Download from pixeldrain.com using their API"""
        try:
            self.log(f"  → Accessing pixeldrain.com...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/octet-stream,*/*'
            }
            
            # Extract file ID from URL
            # Patterns: 
            #   https://pixeldrain.com/u/{id}
            #   https://pixeldrain.com/api/file/{id}
            #   https://pixeldrain.com/l/{list_id}#{file_id}
            file_id = None
            
            # Try /u/{id} pattern (most common)
            match = re.search(r'pixeldrain\.com/u/([a-zA-Z0-9]+)', url)
            if match:
                file_id = match.group(1)
            
            # Try /api/file/{id} pattern
            if not file_id:
                match = re.search(r'pixeldrain\.com/api/file/([a-zA-Z0-9]+)', url)
                if match:
                    file_id = match.group(1)
            
            # Try list pattern with hash
            if not file_id:
                match = re.search(r'pixeldrain\.com/l/[^#]+#([a-zA-Z0-9]+)', url)
                if match:
                    file_id = match.group(1)
            
            if not file_id:
                self.log(f"  ✗ Could not extract file ID from pixeldrain URL")
                return False
            
            self.log(f"  → File ID: {file_id}")
            
            # First, get file info to get the original filename
            info_url = f"https://pixeldrain.com/api/file/{file_id}/info"
            info_response = requests.get(info_url, headers=headers, timeout=15)
            
            original_filename = None
            if info_response.status_code == 200:
                try:
                    info = info_response.json()
                    original_filename = info.get('name', '')
                    file_size = info.get('size', 0)
                    if file_size:
                        size_kb = file_size / 1024
                        if size_kb > 1024:
                            self.log(f"  → File: {original_filename} ({size_kb/1024:.1f} MB)")
                        else:
                            self.log(f"  → File: {original_filename} ({size_kb:.0f} KB)")
                except:
                    pass
            
            # Download the actual file using the API endpoint
            download_url = f"https://pixeldrain.com/api/file/{file_id}"
            self.log(f"  → Downloading...")
            
            response = requests.get(download_url, headers=headers, stream=True, timeout=60)
            response.raise_for_status()
            
            # Check if we got HTML instead of a file (error page)
            content_type = response.headers.get('Content-Type', '')
            if 'text/html' in content_type.lower():
                self.log(f"  ✗ Received HTML instead of file - link may be invalid or expired")
                return False
            
            # Determine extension
            extension = '.mp3'  # Default
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            elif content_type:
                extension = self.infer_extension(content_type)
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Save the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            # Verify the file isn't corrupted (basic check)
            file_size = filepath.stat().st_size
            if file_size < 1000:  # Less than 1KB is suspicious
                self.log(f"  ⚠ Warning: File is very small ({file_size} bytes) - may be corrupted")
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ pixeldrain error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False
    
    def download_fileditch(self, url, output_path, artist, title):
        """Download from fileditch/fileditchfiles"""
        try:
            headers = self.default_headers.copy()
            
            # If it's already a direct files.fileditch.st URL, download directly
            if 'files.fileditch.st' in url or 'files.fileditch.ch' in url:
                download_url = url
                original_filename = Path(urlparse(url).path).name.split('?')[0]
            else:
                # Fetch the page to find the download link
                self.log("  → Fetching fileditch page...")
                response = requests.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Find the download link - it's in an <a> tag with href containing files.fileditch
                download_link = None
                for a in soup.find_all('a', href=True):
                    href = a['href']
                    if 'files.fileditch' in href:
                        download_link = href
                        break
                
                if not download_link:
                    self.log("  ✗ Could not find download link on fileditch page")
                    return False
                
                download_url = download_link
                # Extract original filename from the path
                path_part = urlparse(download_url).path
                original_filename = Path(path_part).name
            
            self.log(f"  → Downloading from fileditch...")
            
            response = requests.get(download_url, headers=headers, stream=True, timeout=60)
            response.raise_for_status()
            
            # Check if we got HTML instead of a file
            content_type = response.headers.get('Content-Type', '')
            if 'text/html' in content_type.lower():
                self.log(f"  ✗ Received HTML instead of file - link may be expired")
                return False
            
            # Determine extension
            extension = '.mp3'  # Default
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            elif content_type:
                extension = self.infer_extension(content_type)
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Save the file
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ fileditch error: {str(e)}")
            return False
    
    def download_bumpworthy(self, url, output_path, artist, title):
        """Download from bumpworthy.com"""
        try:
            import re
            headers = self.default_headers.copy()
            
            # Extract bump ID from URL
            # Supports: /bumps/5215, /download/video/5215, /download/audio/5215
            match = re.search(r'bumpworthy\.com/(?:bumps|download/(?:video|audio))/(\d+)', url)
            if not match:
                self.log("  ✗ Could not extract bump ID from URL")
                return False
            
            bump_id = match.group(1)
            
            # Determine if user wants video or audio based on URL or default to video
            if '/download/audio/' in url:
                download_url = f"https://www.bumpworthy.com/download/audio/{bump_id}"
                default_ext = '.mp3'
            else:
                download_url = f"https://www.bumpworthy.com/download/video/{bump_id}"
                default_ext = '.mp4'
            
            self.log(f"  → Downloading from BumpWorthy (ID: {bump_id})...")
            
            response = requests.get(download_url, headers=headers, stream=True, timeout=60)
            response.raise_for_status()
            
            # Check content type
            content_type = response.headers.get('Content-Type', '')
            if 'text/html' in content_type.lower():
                self.log(f"  ✗ Received HTML - bump may not exist")
                return False
            
            # Determine extension
            extension = self.infer_extension(content_type) if content_type else default_ext
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ bumpworthy error: {str(e)}")
            return False
    
    def download_google_drive(self, url, output_path, artist, title):
        """Download from Google Drive"""
        try:
            import re
            headers = self.default_headers.copy()
            
            # Extract file ID from various Google Drive URL formats
            # /file/d/FILE_ID/view, /open?id=FILE_ID, etc.
            match = re.search(r'(?:/d/|id=|/file/d/)([a-zA-Z0-9_-]+)', url)
            if not match:
                self.log("  ✗ Could not extract Google Drive file ID from URL")
                return False
            
            file_id = match.group(1)
            self.log(f"  → Downloading from Google Drive (ID: {file_id[:16]}...)...")
            
            # Create a session to handle cookies
            session = requests.Session()
            session.headers.update(headers)
            
            # Try multiple download methods
            download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
            
            response = session.get(download_url, stream=True, timeout=60)
            
            # Check for virus scan warning page (large files > 100MB)
            content_type = response.headers.get('Content-Type', '')
            if 'text/html' in content_type:
                self.log("  → Large file detected, getting confirmation...")
                html_content = response.text
                
                # Method 1: Look for download_warning cookie
                confirm_token = None
                for key, value in session.cookies.items():
                    if 'download_warning' in key:
                        confirm_token = value
                        break
                
                # Method 2: Look for confirm token in HTML using various patterns
                if not confirm_token:
                    patterns = [
                        r'confirm=([0-9A-Za-z_-]+)',
                        r'name="confirm" value="([^"]+)"',
                        r'/uc\?export=download&amp;confirm=([^&]+)',
                        r'download&amp;confirm=([^&"]+)',
                    ]
                    for pattern in patterns:
                        match = re.search(pattern, html_content)
                        if match:
                            confirm_token = match.group(1)
                            break
                
                # Method 3: Try with confirm=t (works for some files)
                if not confirm_token:
                    confirm_token = 't'
                
                # Retry with confirmation
                download_url = f"https://drive.google.com/uc?export=download&confirm={confirm_token}&id={file_id}"
                response = session.get(download_url, stream=True, timeout=60)
                content_type = response.headers.get('Content-Type', '')
                
                # If still HTML, try the direct download URL format
                if 'text/html' in content_type:
                    self.log("  → Trying alternate download method...")
                    # Try the drive.usercontent.google.com endpoint
                    download_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
                    response = session.get(download_url, stream=True, timeout=60)
                    content_type = response.headers.get('Content-Type', '')
            
            response.raise_for_status()
            
            # Final check if we still got HTML
            if 'text/html' in content_type.lower():
                self.log("  ✗ Could not download file - may require login or be restricted")
                return False
            
            # Try to get filename from Content-Disposition header
            content_disp = response.headers.get('Content-Disposition', '')
            original_filename = None
            if 'filename' in content_disp:
                # Handle both filename= and filename*=UTF-8''
                fname_match = re.search(r"filename\*?=(?:UTF-8'')?\"?([^\";\n]+)\"?", content_disp, re.IGNORECASE)
                if fname_match:
                    original_filename = fname_match.group(1)
                    from urllib.parse import unquote
                    original_filename = unquote(original_filename).strip('"')
                    self.log(f"  → Original filename: {original_filename}")
            
            # Determine extension from original filename or content type
            extension = '.mp3'  # Default
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            elif content_type:
                extension = self.infer_extension(content_type)
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Download with progress indication
            total_size = response.headers.get('Content-Length')
            if total_size:
                self.log(f"  → Downloading ({int(total_size)//1024} KB)...")
            
            downloaded = 0
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
                    downloaded += len(chunk)
            
            final_size = filepath.stat().st_size
            if final_size < 100:
                self.log(f"  ⚠ Warning: File is very small ({final_size} bytes)")
            
            self.log(f"  ✓ Saved as: {filepath.name} ({final_size//1024} KB)")
            return True
            
        except Exception as e:
            self.log(f"  ✗ Google Drive error: {str(e)}")
            return False
    
    def download_mega(self, url, output_path, artist, title):
        """Download from MEGA.nz"""
        try:
            import base64
            import struct
            import json as json_lib
            from Crypto.Cipher import AES
            from Crypto.Util import Counter
            import re
            import random
            
            self.log("  → Parsing MEGA link...")
            
            # Parse MEGA URL to extract file ID and key
            # Formats: mega.nz/file/FILEID#KEY or mega.nz/#!FILEID!KEY (old format)
            match = re.search(r'mega\.(?:nz|co\.nz)(?:/file/|/#!|/folder/)([^#!]+)[#!](.+?)(?:\?|$)', url)
            if not match:
                self.log("  ✗ Could not parse MEGA URL")
                return False
            
            file_id = match.group(1)
            file_key = match.group(2)
            
            # Decode the key (base64url to bytes)
            def base64_url_decode(data):
                data = data.replace('-', '+').replace('_', '/').replace(',', '')
                padding = 4 - len(data) % 4
                if padding != 4:
                    data += '=' * padding
                return base64.b64decode(data)
            
            # Convert key bytes to array of 32-bit integers
            def str_to_a32(s):
                # Pad to multiple of 4
                if len(s) % 4:
                    s += b'\x00' * (4 - len(s) % 4)
                return struct.unpack('>%dI' % (len(s) // 4), s)
            
            def a32_to_str(a):
                return struct.pack('>%dI' % len(a), *a)
            
            # Decrypt key
            def decrypt_key(key_a32):
                # MEGA uses a compound key - split and XOR
                return (
                    key_a32[0] ^ key_a32[4],
                    key_a32[1] ^ key_a32[5],
                    key_a32[2] ^ key_a32[6],
                    key_a32[3] ^ key_a32[7]
                )
            
            key_bytes = base64_url_decode(file_key)
            key_a32 = str_to_a32(key_bytes)
            
            # Get file key and IV
            if len(key_a32) == 8:
                file_key_a32 = decrypt_key(key_a32)
                iv = (key_a32[4], key_a32[5], 0, 0)
            else:
                self.log("  ✗ Invalid MEGA key format")
                return False
            
            # Call MEGA API to get file info
            self.log("  → Fetching file info from MEGA API...")
            
            api_url = "https://g.api.mega.co.nz/cs"
            seq_no = random.randint(0, 0xFFFFFFFF)
            
            request_data = [{"a": "g", "g": 1, "p": file_id}]
            
            response = requests.post(
                f"{api_url}?id={seq_no}",
                json=request_data,
                timeout=30
            )
            response.raise_for_status()
            
            result = response.json()
            if isinstance(result, int) and result < 0:
                error_msgs = {
                    -1: "Internal error",
                    -2: "Invalid arguments",
                    -3: "Request failed, retrying",
                    -9: "File not found",
                    -11: "Access denied",
                    -14: "Temporarily unavailable",
                    -16: "User blocked",
                    -17: "Request quota exceeded",
                    -18: "Resource unavailable"
                }
                self.log(f"  ✗ MEGA API error: {error_msgs.get(result, f'Error code {result}')}")
                return False
            
            file_info = result[0]
            if isinstance(file_info, int):
                self.log(f"  ✗ MEGA file error: {file_info}")
                return False
            
            download_url = file_info.get('g')
            file_size = file_info.get('s', 0)
            encrypted_attrs = file_info.get('at', '')
            
            if not download_url:
                self.log("  ✗ Could not get download URL from MEGA")
                return False
            
            # Decrypt file attributes to get filename
            original_filename = None
            try:
                attrs_bytes = base64_url_decode(encrypted_attrs)
                cipher = AES.new(a32_to_str(file_key_a32), AES.MODE_CBC, b'\x00' * 16)
                decrypted = cipher.decrypt(attrs_bytes)
                # Remove padding and parse JSON
                decrypted = decrypted.rstrip(b'\x00')
                if decrypted.startswith(b'MEGA'):
                    json_str = decrypted[4:].decode('utf-8', errors='ignore')
                    # Find JSON object
                    json_match = re.search(r'\{.*\}', json_str)
                    if json_match:
                        attrs = json_lib.loads(json_match.group())
                        original_filename = attrs.get('n', '')
                        self.log(f"  → Original filename: {original_filename}")
            except Exception as e:
                self.log(f"  → Could not decrypt filename: {e}")
            
            self.log(f"  → Downloading encrypted file ({file_size // 1024} KB)...")
            
            # Download encrypted file
            response = requests.get(download_url, stream=True, timeout=120)
            response.raise_for_status()
            
            # Decrypt the file as we download
            k_str = a32_to_str(file_key_a32)
            iv_int = struct.unpack('>Q', a32_to_str(iv[:2]))[0]
            ctr = Counter.new(128, initial_value=iv_int << 64)
            cipher = AES.new(k_str, AES.MODE_CTR, counter=ctr)
            
            # Determine extension
            extension = '.mp3'
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            # Download and decrypt
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=16384):
                    if not self.is_downloading:
                        return False
                    decrypted_chunk = cipher.decrypt(chunk)
                    f.write(decrypted_chunk)
            
            # Truncate to actual file size (remove padding)
            with open(filepath, 'r+b') as f:
                f.truncate(file_size)
            
            final_size = filepath.stat().st_size
            self.log(f"  ✓ Saved as: {filepath.name} ({final_size // 1024} KB)")
            return True
            
        except ImportError:
            self.log("  ✗ MEGA download requires pycryptodome. Install with: pip install pycryptodome")
            return False
        except Exception as e:
            self.log(f"  ✗ MEGA error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False
    
    def download_imgur(self, url, output_path, artist, title):
        """Download image from Imgur"""
        try:
            # Handle different Imgur URL formats
            # https://imgur.com/a/XXXXX (album)
            # https://imgur.com/XXXXX (single image page)
            # https://i.imgur.com/XXXXX.jpg (direct image)
            
            parsed = urlparse(url)
            path = parsed.path.strip('/')
            
            # If it's already a direct image URL (i.imgur.com)
            if 'i.imgur.com' in url:
                direct_url = url
                # Ensure it has an extension
                if not any(url.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.mp4', '.webp']):
                    direct_url = url + '.jpg'
            # Album URL - try to get first image
            elif '/a/' in url or '/gallery/' in url:
                album_id = path.split('/')[-1]
                # Try to fetch album page and extract first image
                self.log(f"    Fetching Imgur album: {album_id}")
                response = requests.get(url, headers=self.default_headers, timeout=30)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.text, 'html.parser')
                # Look for image tags
                img_tag = soup.find('img', class_='post-image-placeholder') or soup.find('img', src=lambda x: x and 'i.imgur.com' in x)
                if img_tag and img_tag.get('src'):
                    direct_url = img_tag['src']
                    if direct_url.startswith('//'):
                        direct_url = 'https:' + direct_url
                else:
                    # Try finding in meta tags
                    meta = soup.find('meta', property='og:image')
                    if meta and meta.get('content'):
                        direct_url = meta['content']
                    else:
                        self.log(f"    Could not find image in album")
                        return False
            # Single image page
            else:
                image_id = path.split('/')[-1].split('.')[0]
                # Try common extensions
                direct_url = f"https://i.imgur.com/{image_id}.jpg"
            
            self.log(f"    Direct URL: {direct_url}")
            
            # Download the image
            response = requests.get(direct_url, headers=self.default_headers, stream=True, timeout=30)
            response.raise_for_status()
            
            # Get extension from URL or content type
            extension = self.infer_image_extension(response.headers.get('Content-Type'), direct_url)
            filename = f"{self.build_safe_title(title)}{extension}"
            filepath = self.resolve_duplicate_path(Path(output_path) / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"    Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"    Imgur error: {str(e)}")
            return False
    
    def download_ibb(self, url, output_path, artist, title):
        """Download image from ibb.co (ImgBB)"""
        try:
            # Handle different ibb.co URL formats
            # https://ibb.co/XXXXX (page URL)
            # https://i.ibb.co/XXXXX/filename.jpg (direct image)
            
            # If it's already a direct image URL (i.ibb.co)
            if 'i.ibb.co' in url:
                direct_url = url
            else:
                # Fetch the page and extract the direct image URL
                self.log(f"    Fetching ibb.co page...")
                response = requests.get(url, headers=self.default_headers, timeout=30)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.text, 'html.parser')
                
                # Look for direct image in img tags with i.ibb.co
                direct_url = None
                for img in soup.find_all('img'):
                    src = img.get('src', '')
                    if 'i.ibb.co' in src:
                        direct_url = src
                        break
                
                # Fallback to og:image meta tag
                if not direct_url:
                    meta = soup.find('meta', property='og:image')
                    if meta and meta.get('content'):
                        direct_url = meta['content']
                
                if not direct_url:
                    self.log(f"    Could not find image URL on ibb.co page")
                    return False
            
            self.log(f"    Direct URL: {direct_url}")
            
            # Download the image
            response = requests.get(direct_url, headers=self.default_headers, stream=True, timeout=30)
            response.raise_for_status()
            
            # Get extension from URL or content type
            extension = self.infer_image_extension(response.headers.get('Content-Type'), direct_url)
            filename = f"{self.build_safe_title(title)}{extension}"
            filepath = self.resolve_duplicate_path(Path(output_path) / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"    Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"    ibb.co error: {str(e)}")
            return False

    def download_gofile(self, url, output_path, artist, title):
        """Download from gofile.io using their API"""
        try:
            self.log(f"  → Accessing gofile.io...")
            
            # Extract content ID from URL
            # Format: https://gofile.io/d/{contentId}
            match = re.search(r'gofile\.io/d/([a-zA-Z0-9]+)', url)
            if not match:
                self.log(f"  ✗ Could not extract content ID from gofile URL")
                return False
            
            content_id = match.group(1)
            self.log(f"  → Content ID: {content_id}")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/json',
                'Origin': 'https://gofile.io',
                'Referer': 'https://gofile.io/'
            }
            
            # First, create a guest account to get a token
            self.log(f"  → Getting access token...")
            account_response = requests.post(
                'https://api.gofile.io/accounts',
                headers=headers,
                timeout=30
            )
            
            if account_response.status_code != 200:
                self.log(f"  ✗ Failed to create guest account")
                return False
            
            account_data = account_response.json()
            if account_data.get('status') != 'ok':
                self.log(f"  ✗ Account creation failed: {account_data.get('status')}")
                return False
            
            token = account_data.get('data', {}).get('token')
            if not token:
                self.log(f"  ✗ No token received from gofile")
                return False
            
            # Now get the content info with the token
            headers['Authorization'] = f'Bearer {token}'
            headers['Cookie'] = f'accountToken={token}'
            
            content_url = f'https://api.gofile.io/contents/{content_id}?wt=4fd6sg89d7s6'
            self.log(f"  → Fetching content info...")
            
            content_response = requests.get(content_url, headers=headers, timeout=30)
            
            if content_response.status_code != 200:
                self.log(f"  ✗ Failed to get content info (status {content_response.status_code})")
                return False
            
            content_data = content_response.json()
            
            if content_data.get('status') != 'ok':
                error_msg = content_data.get('status', 'Unknown error')
                self.log(f"  ✗ Content request failed: {error_msg}")
                # Check for password protection
                if 'password' in str(error_msg).lower() or content_data.get('data', {}).get('passwordStatus') == 'passwordRequired':
                    self.log(f"  ✗ This content is password protected")
                return False
            
            # Extract files from the content
            data = content_data.get('data', {})
            children = data.get('children', {})
            
            if not children:
                # Maybe it's a direct file, not a folder
                if data.get('type') == 'file':
                    children = {content_id: data}
                else:
                    self.log(f"  ✗ No files found in gofile content")
                    return False
            
            # Download each file
            downloaded_any = False
            for file_id, file_info in children.items():
                if file_info.get('type') != 'file':
                    continue
                
                file_name = file_info.get('name', 'unknown')
                file_link = file_info.get('link')
                file_size = file_info.get('size', 0)
                
                if not file_link:
                    self.log(f"  ✗ No download link for {file_name}")
                    continue
                
                # Format file size
                if file_size > 1024 * 1024:
                    size_str = f"{file_size / (1024*1024):.1f} MB"
                elif file_size > 1024:
                    size_str = f"{file_size / 1024:.0f} KB"
                else:
                    size_str = f"{file_size} bytes"
                
                self.log(f"  → Downloading: {file_name} ({size_str})")
                
                # Download the file
                download_headers = headers.copy()
                download_headers['Accept'] = '*/*'
                
                try:
                    download_response = requests.get(
                        file_link,
                        headers=download_headers,
                        stream=True,
                        timeout=120
                    )
                    download_response.raise_for_status()
                    
                    # Get extension from original filename
                    extension = Path(file_name).suffix or '.mp3'
                    
                    # Build output filename
                    if len(children) == 1:
                        # Single file - use provided title
                        filename = self.build_track_filename(title, extension=extension)
                    else:
                        # Multiple files - use original filename
                        safe_name = self.build_safe_title(Path(file_name).stem)
                        filename = f"{safe_name}{extension}"
                    
                    filepath = self.resolve_duplicate_path(Path(output_path) / filename)
                    
                    with open(filepath, 'wb') as f:
                        for chunk in download_response.iter_content(chunk_size=8192):
                            if not self.is_downloading:
                                return downloaded_any
                            f.write(chunk)
                    
                    self.log(f"  ✓ Saved as: {filepath.name}")
                    downloaded_any = True
                    
                except Exception as e:
                    self.log(f"  ✗ Failed to download {file_name}: {str(e)}")
                    continue
            
            return downloaded_any
            
        except Exception as e:
            self.log(f"  ✗ gofile error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False

    def download_mediafire(self, url, output_path, artist, title):
        """Download from mediafire.com"""
        try:
            self.log(f"  → Accessing MediaFire page...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            }
            
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Get download link from download button
            download_btn = soup.find('a', {'id': 'downloadButton'})
            if not download_btn:
                download_btn = soup.find('a', {'aria-label': 'Download file'})
            
            direct_url = None
            if download_btn:
                direct_url = download_btn.get('href')
            
            # Fallback: search for direct URL pattern in page
            if not direct_url:
                match = re.search(r'https://download\d*\.mediafire\.com/[^"\'<>\s]+', response.text)
                if match:
                    direct_url = match.group(0)
            
            if not direct_url:
                self.log(f"  ✗ Could not find download link on MediaFire page")
                return False
            
            self.log(f"  → Found direct download link")
            
            # Get original filename from page
            original_filename = None
            filename_div = soup.find('div', class_='filename')
            if filename_div:
                original_filename = filename_div.text.strip()
            
            # Fallback: extract from URL
            if not original_filename:
                original_filename = direct_url.split('/')[-1]
            
            if original_filename:
                self.log(f"  → File: {original_filename}")
            
            # Download the file
            self.log(f"  → Downloading...")
            dl_headers = headers.copy()
            dl_headers['Referer'] = url
            
            dl_response = requests.get(direct_url, headers=dl_headers, stream=True, timeout=120)
            dl_response.raise_for_status()
            
            # Check we didn't get an error page
            content_type = dl_response.headers.get('Content-Type', '')
            if 'text/html' in content_type.lower():
                self.log(f"  ✗ Received HTML instead of file - link may be invalid or expired")
                return False
            
            # Determine extension
            extension = '.mp3'  # Default
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            elif content_type:
                extension = self.infer_extension(content_type)
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(Path(output_path) / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in dl_response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ MediaFire error: {str(e)}")
            import traceback
            self.log(f"  Debug: {traceback.format_exc()}")
            return False

    def download_aws_s3(self, url, output_path, artist, title):
        """Download from AWS S3 (public buckets)"""
        try:
            self.log(f"  → Downloading from AWS S3...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            }
            
            # Extract filename from URL path
            from urllib.parse import urlparse, unquote
            parsed = urlparse(url)
            path_parts = parsed.path.split('/')
            original_filename = unquote(path_parts[-1]) if path_parts else None
            
            if original_filename:
                self.log(f"  → File: {original_filename}")
            
            # Try HEAD request to get content info
            try:
                head = requests.head(url, headers=headers, timeout=15, allow_redirects=True)
                if head.status_code == 200:
                    content_length = head.headers.get('Content-Length')
                    if content_length:
                        size = int(content_length)
                        if size > 1024 * 1024:
                            self.log(f"  → Size: {size / (1024*1024):.1f} MB")
                        else:
                            self.log(f"  → Size: {size / 1024:.0f} KB")
            except:
                pass
            
            # Download the file
            response = requests.get(url, headers=headers, stream=True, timeout=120)
            response.raise_for_status()
            
            # Check content type
            content_type = response.headers.get('Content-Type', '')
            
            # Determine extension
            extension = '.mp3'  # Default
            if original_filename:
                ext = Path(original_filename).suffix
                if ext:
                    extension = ext
            elif content_type:
                extension = self.infer_extension(content_type)
            
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(Path(output_path) / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  ✗ AWS S3 error: {str(e)}")
            return False

    def download_imgurgg(self, url, output_path, artist, title):
        """Download audio/video from imgur.gg"""
        try:
            # Handle different imgur.gg URL formats
            # https://imgur.gg/f/XXXXX (file page)
            # https://i.imgur.gg/XXXXX-filename.mp3 (direct file)
            
            # If it's already a direct file URL (i.imgur.gg)
            if 'i.imgur.gg' in url:
                direct_url = url
            else:
                # Fetch the page to find the media source
                self.log(f"    Fetching imgur.gg page...")
                response = requests.get(url, headers=self.default_headers, timeout=30)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.text, 'html.parser')
                direct_url = None
                
                # First, check og:video or og:audio meta tags (most reliable)
                og_video = soup.find('meta', property='og:video')
                og_audio = soup.find('meta', property='og:audio')
                twitter_player = soup.find('meta', attrs={'name': 'twitter:player:stream'})
                
                if og_video and og_video.get('content'):
                    direct_url = og_video['content']
                    self.log(f"    Found video in og:video meta tag")
                elif og_audio and og_audio.get('content'):
                    direct_url = og_audio['content']
                    self.log(f"    Found audio in og:audio meta tag")
                elif twitter_player and twitter_player.get('content'):
                    direct_url = twitter_player['content']
                    self.log(f"    Found media in twitter:player:stream meta tag")
                else:
                    # Fallback: Look for audio or video source tags
                    source = soup.find('audio', {'src': True}) or soup.find('video', {'src': True})
                    if source and source.get('src'):
                        direct_url = source['src']
                    else:
                        # Try finding source tag inside audio/video
                        source_tag = soup.select_one('audio source[src], video source[src]')
                        if source_tag and source_tag.get('src'):
                            direct_url = source_tag['src']
                
                if not direct_url:
                    self.log(f"    Could not find media source on imgur.gg page")
                    return False
                
                # Handle relative URLs
                if direct_url.startswith('//'):
                    direct_url = 'https:' + direct_url
                elif direct_url.startswith('/'):
                    direct_url = 'https://imgur.gg' + direct_url
            
            self.log(f"    Direct URL: {direct_url}")
            
            # Download the file
            response = requests.get(direct_url, headers=self.default_headers, stream=True, timeout=60)
            response.raise_for_status()
            
            # Get extension from URL
            parsed_url = urlparse(direct_url)
            url_path = parsed_url.path
            if '.' in url_path.split('/')[-1]:
                extension = '.' + url_path.split('.')[-1].lower()
            else:
                # Infer from content type
                content_type = response.headers.get('Content-Type', '')
                if 'audio/mpeg' in content_type:
                    extension = '.mp3'
                elif 'audio/mp4' in content_type or 'audio/m4a' in content_type:
                    extension = '.m4a'
                elif 'video/mp4' in content_type:
                    extension = '.mp4'
                else:
                    extension = '.mp3'  # Default to mp3 for imgur.gg audio
            
            filename = f"{self.build_safe_title(title)}{extension}"
            filepath = self.resolve_duplicate_path(Path(output_path) / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
            
            self.log(f"    Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"    imgur.gg error: {str(e)}")
            return False
            
    def download_direct(self, url, output_path, artist, title):
        """Download from direct URL"""
        try:
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            
            # Get file extension from URL or content-type
            extension = Path(urlparse(url).path).suffix
            if not extension:
                extension = self.infer_extension(response.headers.get('Content-Type'))
            filename = self.build_track_filename(title, extension=extension)
            filepath = self.resolve_duplicate_path(output_path / filename)
            
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if not self.is_downloading:
                        return False
                    f.write(chunk)
                    
            self.log(f"  ✓ Saved as: {filepath.name}")
            return True
            
        except Exception as e:
            self.log(f"  Download error: {str(e)}")
            return False
            
    def create_zip_archive(self):
        """Create a ZIP archive of all downloaded files"""
        try:
            output_folder = Path(self.output_folder_var.get())
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = output_folder / f"music_archive_{timestamp}.zip"
            
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_folder):
                    for file in files:
                        if file.endswith('.zip'):
                            continue
                        filepath = Path(root) / file
                        arcname = filepath.relative_to(output_folder)
                        zipf.write(filepath, arcname)
                        
            self.log(f"✓ ZIP archive created: {zip_filename.name}")
            
        except Exception as e:
            self.log(f"✗ Failed to create ZIP: {str(e)}")


def main():
    root = tk.Tk()
    app = MusicDownloaderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
