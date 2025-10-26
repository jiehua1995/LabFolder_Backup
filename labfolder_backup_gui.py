#!/usr/bin/env python3
# labfolder_backup_v6.py
# Tkinter GUI wrapper for labfolder_backup_v4.LabFolderBackup

import os
import sys
import threading
import queue
from pathlib import Path
import tkinter as tk
from tkinter import Tk, LEFT
from tkinter import font as tkfont
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import webbrowser

# Debugging: Print current working directory and sys.path
print(f"Current working directory: {os.getcwd()}")
print(f"sys.path: {sys.path}")

# Handle PyInstaller runtime environment
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    BASE_DIR = sys._MEIPASS  # For loading bundled resources
    # For saving .env, use the directory where the .exe is located
    EXECUTABLE_DIR = Path(sys.executable).parent
else:
    # Running in a normal Python environment
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXECUTABLE_DIR = Path(BASE_DIR)

# Update paths for bundled files
ENV_FILE = Path(BASE_DIR) / '.env'  # For loading
ENV_SAVE_FILE = EXECUTABLE_DIR / '.env'  # For saving
V4_SCRIPT = Path(BASE_DIR) / 'labfolder_backup_v4.py'

# Ensure .env exists
if not ENV_FILE.exists():
    print("Error: .env file not found.")
    sys.exit(1)

# Ensure labfolder_backup_v4.py exists
if not V4_SCRIPT.exists():
    print("Error: labfolder_backup_v4.py not found.")
    sys.exit(1)

# Ensure labfolder_backup_v4 is importable
try:
    from labfolder_backup_v4 import LabFolderBackup
except ImportError as e:
    print(f"Error importing LabFolderBackup: {e}")
    sys.exit(1)

class QueueWriter:
    """Write wrapper that pushes (source, text) tuples into a queue and
    forwards the write to the original stream.
    source should be 'out' or 'err'.
    """

    def __init__(self, q, orig, source='out'):
        self.q = q
        self.orig = orig
        self.source = source

    def write(self, s):
        try:
            if s:
                self.q.put((self.source, s))
        except Exception:
            pass
        try:
            self.orig.write(s)
        except Exception:
            pass

    def flush(self):
        try:
            self.orig.flush()
        except Exception:
            pass


def load_env_defaults():
    defaults = {}
    if ENV_FILE.exists():
        try:
            with open(ENV_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#') or '=' not in line:
                        continue
                    k, v = line.split('=', 1)
                    defaults[k.strip()] = v.strip()
        except Exception:
            pass

    def g(k, fallback=''):
        return defaults.get(k, os.getenv(k, fallback))

    return {
        'LABFOLDER_URL': g('LABFOLDER_URL', ''),
        'LABFOLDER_USERNAME': g('LABFOLDER_USERNAME', ''),
        'LABFOLDER_PASSWORD': g('LABFOLDER_PASSWORD', ''),
        'DOWNLOAD_DIR': g('DOWNLOAD_DIR', 'labfolder_backup'),
        'SCROLL_NUDGE_PIXELS': g('SCROLL_NUDGE_PIXELS', '250'),
    }


class BackupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('LabFolder Backup (GUI)')
        # Try to set a friendly light background color for a modern look
        try:
            self.root.configure(bg='#f7f9fc')
        except Exception:
            pass

        self.q = queue.Queue()
        self.worker_thread = None
        self.running = False

        defaults = load_env_defaults()

        # Fonts and styles
        self.title_font = tkfont.Font(family='Segoe UI', size=14, weight='bold')
        self.subtitle_font = tkfont.Font(family='Segoe UI', size=9)
        self.label_font = ('Segoe UI', 10)
        self.entry_font = ('Consolas', 10)

        # Modern light theme (cross-platform, no theme switching)
        try:
            style = ttk.Style()
            # Prefer 'clam' for consistent cross-platform look; fallback to current
            try:
                style.theme_use('clam')
            except Exception:
                pass

            # Palette
            bg = '#f7f9fc'
            panel = '#ffffff'
            accent = '#2b6ea3'  # button/primary accent
            muted = '#6b7280'
            success = '#059669'

            style.configure('TFrame', background=bg)
            style.configure('TLabel', background=bg, foreground='#111111')
            style.configure('Muted.TLabel', background=bg, foreground=muted)
            style.configure('TButton', padding=6, relief='flat')
            style.map('TButton', background=[('active', accent), ('!disabled', accent)], foreground=[('!disabled', '#ffffff')])
            style.configure('Accent.TButton', background=accent, foreground='#ffffff')
            style.configure('TEntry', fieldbackground=panel, background=panel)
            style.configure('TSeparator', background='#e6edf3')
        except Exception:
            style = None

        # Header: big plain title (no colored background)
        self.title_font = tkfont.Font(family='Segoe UI', size=20, weight='bold')
        header = ttk.Frame(self.root)
        header.pack(fill='x', padx=8, pady=(8, 0))
        hdr = ttk.Label(header, text='LabFolder Backup', style='TLabel')
        hdr.configure(font=self.title_font)
        hdr.pack(side='top', anchor='w', padx=10, pady=(0, 4))
        # optional subtitle (smaller, muted)
        sub = ttk.Label(header, text='Use at your own risk. Click "Run Backup" to start', style='Muted.TLabel')
        sub.configure(font=self.subtitle_font)
        sub.pack(side='top', anchor='w', padx=10, pady=(0, 4))

        # Add copyright text below the title
        copyright_label = ttk.Label(header, text='Developed by Imhof Group, BMC, LMU Munich.', style='Muted.TLabel')
        copyright_label.configure(font=self.subtitle_font)
        copyright_label.pack(side='top', anchor='w', padx=10, pady=(0, 12))

        # Main top area
        top = ttk.Frame(self.root, padding=(12, 8, 12, 8))
        top.pack(fill='x')

        # Variables
        self.url_var = tk.StringVar(value=defaults['LABFOLDER_URL'])
        self.user_var = tk.StringVar(value=defaults['LABFOLDER_USERNAME'])
        self.pwd_var = tk.StringVar(value=defaults['LABFOLDER_PASSWORD'])
        self.dir_var = tk.StringVar(value=defaults['DOWNLOAD_DIR'])
        self.nudge_var = tk.StringVar(value=defaults['SCROLL_NUDGE_PIXELS'])

        # Row 1: URL
        row = ttk.Frame(top)
        row.pack(fill='x', pady=4)
        lbl = ttk.Label(row, text='LabFolder URL:')
        lbl.pack(side=LEFT)
        self.url_entry = ttk.Entry(row, textvariable=self.url_var, width=60)
        self.url_entry.pack(side=LEFT, padx=6)

        # Row 2: username / password
        row = ttk.Frame(top)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text='Username:').pack(side=LEFT)
        ttk.Entry(row, textvariable=self.user_var, width=20).pack(side=LEFT, padx=6)
        ttk.Label(row, text='Password:').pack(side=LEFT)
        ttk.Entry(row, textvariable=self.pwd_var, width=20, show='*').pack(side=LEFT, padx=6)

        # Row 3: download dir
        row = ttk.Frame(top)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text='Download Dir:').pack(side=LEFT)
        ttk.Entry(row, textvariable=self.dir_var, width=40).pack(side=LEFT, padx=6)
        ttk.Button(row, text='Open', command=self.open_download_dir).pack(side=LEFT, padx=6)

        # Row 4: nudge
        row = ttk.Frame(top)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text='Scroll Nudge (px):').pack(side=LEFT)
        ttk.Entry(row, textvariable=self.nudge_var, width=6).pack(side=LEFT, padx=6)

        # Buttons
        btn_row = ttk.Frame(top)
        btn_row.pack(fill='x', pady=8)
        self.run_btn = ttk.Button(btn_row, text='Run Backup', command=self.start_backup)
        self.run_btn.pack(side=LEFT, padx=6)
        ttk.Button(btn_row, text='Save .env', command=self.save_env).pack(side=LEFT, padx=6)
        ttk.Button(btn_row, text='Save Log', command=self.save_log).pack(side=LEFT, padx=6)
        ttk.Button(btn_row, text='Check Output', command=self.open_index).pack(side=LEFT, padx=6)

        # Scrolled text for logs (clean white panel)
        self.log = ScrolledText(self.root, height=20, wrap='word', bg='#ffffff', relief='solid', bd=1)
        self.log.pack(fill='both', expand=True, padx=12, pady=(8, 6))

        # Configure font and tags for colorblind-friendly palette
        try:
            self.log.configure(font=('Consolas', 11), foreground='#0f1724')
        except Exception:
            pass
        # Color-blind friendly palette (blue, green, yellow, orange, magenta)
        self.log.tag_config('info', foreground='#0072B2')
        self.log.tag_config('success', foreground='#009E73')
        self.log.tag_config('warning', foreground='#F0E442')
        self.log.tag_config('error', foreground='#D55E00')
        self.log.tag_config('stderr', foreground='#CC79A7')

        # Install queue writer for stdout/stderr
        self.orig_stdout = sys.stdout
        self.orig_stderr = sys.stderr
        sys.stdout = QueueWriter(self.q, self.orig_stdout, source='out')
        sys.stderr = QueueWriter(self.q, self.orig_stderr, source='err')

        # Poll queue
        self._poll_queue()

        # Footer (centered, subtle)
        try:
            sep = ttk.Separator(self.root, orient='horizontal')
            sep.pack(fill='x', padx=10, pady=(6, 4))
            footer = ttk.Frame(self.root)
            footer.pack(fill='x', pady=(4, 12))
            self.footer_font = tkfont.Font(family='Segoe UI', size=9)
            foot_label = ttk.Label(footer, text='Powered by Jie Hua · Copyright Imhof Group, BMC, LMU Munich.', style='Muted.TLabel')
            foot_label.configure(font=self.footer_font)
            foot_label.pack(side='top', pady=2)
        except Exception:
            pass

        # Close handling
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _poll_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                src = None
                text = ''
                if isinstance(item, tuple) and len(item) == 2:
                    src, text = item
                else:
                    text = str(item)

                for line in text.splitlines(True):
                    tag = 'info'
                    l = line.strip()
                    if src == 'err':
                        tag = 'error'
                    else:
                        if any(ch in l for ch in ('✅', '✓')):
                            tag = 'success'
                        elif any(ch in l for ch in ('⚠', '⚠️')):
                            tag = 'warning'
                        elif any(ch in l for ch in ('✗', '✖', '❌', '✕')):
                            tag = 'error'
                        else:
                            tag = 'info'

                    try:
                        self.log.insert('end', line, (tag,))
                        self.log.see('end')
                    except Exception:
                        try:
                            self.log.insert('end', line)
                            self.log.see('end')
                        except Exception:
                            pass
        except queue.Empty:
            pass
        self.root.after(200, self._poll_queue)

    def open_download_dir(self):
        p = Path(self.dir_var.get() or 'labfolder_backup').absolute()
        if not p.exists():
            try:
                p.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"Could not create download dir: {e}")
                return
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(p))
            else:
                webbrowser.open(str(p))
        except Exception as e:
            print(f"Could not open folder: {e}")

    def open_index(self):
        # Open the download directory instead of trying to open index.html
        p = Path(self.dir_var.get() or 'labfolder_backup').absolute()
        if not p.exists():
            try:
                p.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                print(f"Could not create/open download dir: {e}")
                return
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(p))
            else:
                webbrowser.open(str(p))
        except Exception as e:
            print(f"Could not open download folder: {e}")

    def save_env(self):
        """Save .env to the same directory as the script/executable"""
        try:
            lines = [
                f"LABFOLDER_URL={self.url_var.get()}",
                f"LABFOLDER_USERNAME={self.user_var.get()}",
                f"LABFOLDER_PASSWORD={self.pwd_var.get()}",
                f"DOWNLOAD_DIR={self.dir_var.get()}",
                f"SCROLL_NUDGE_PIXELS={self.nudge_var.get()}"
            ]
            with open(ENV_SAVE_FILE, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines) + '\n')
            print(f"✅ Saved .env to {ENV_SAVE_FILE.absolute()}")
        except Exception as e:
            print(f"❌ Failed to save .env: {e}")

    def save_log(self):
        """Save the log window content to a timestamped log file in the same directory"""
        try:
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = EXECUTABLE_DIR / f'backup_log_{timestamp}.txt'
            
            # Get all text from the log widget
            log_content = self.log.get('1.0', 'end-1c')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(log_content)
            
            print(f"✅ Saved log to {log_file.absolute()}")
        except Exception as e:
            print(f"❌ Failed to save log: {e}")

    def start_backup(self):
        if self.running:
            print('Backup already running')
            return

        if LabFolderBackup is None:
            print('LabFolderBackup implementation not available (import error). Check labfolder_backup_v4.py')
            return

        # Validate some inputs
        url = self.url_var.get().strip()
        user = self.user_var.get().strip()
        pwd = self.pwd_var.get().strip()
        if not url or not user or not pwd:
            print('Please provide URL, username and password')
            return

        # Persist env to os.environ for the run
        os.environ['LABFOLDER_URL'] = url
        os.environ['LABFOLDER_USERNAME'] = user
        os.environ['LABFOLDER_PASSWORD'] = pwd
        os.environ['DOWNLOAD_DIR'] = self.dir_var.get().strip() or 'labfolder_backup'
        os.environ['SCROLL_NUDGE_PIXELS'] = self.nudge_var.get().strip() or '250'

        try:
            self.run_btn.config(state='disabled')
        except Exception:
            pass
        self.running = True
        self.log.insert('end', '\n=== Starting backup ===\n')
        self.log.see('end')

        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()

        # Poll thread completion
        self.root.after(1000, self._check_worker)

    def _run_worker(self):
        try:
            backup = LabFolderBackup()
            backup.out = Path(os.environ.get('DOWNLOAD_DIR', 'labfolder_backup'))
            success = backup.run()
            if success:
                print('\n=== Backup finished successfully ===\n')
            else:
                print('\n=== Backup finished with errors or was aborted ===\n')
        except Exception as e:
            print(f"Worker error: {e}")
        finally:
            self.running = False

    def _check_worker(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(1000, self._check_worker)
        else:
            try:
                self.run_btn.config(state='normal')
            except Exception:
                pass
            if not self.running:
                self.log.insert('end', '\n=== Worker stopped ===\n')
                self.log.see('end')

    def _on_close(self):
        # restore stdout/stderr
        try:
            sys.stdout = self.orig_stdout
            sys.stderr = self.orig_stderr
        except Exception:
            pass
        self.root.destroy()


def main():
    root = Tk()
    # Try to set window icon from image.ico in this folder
    try:
        ico = Path(__file__).parent / 'image.ico'
        if ico.exists():
            root.iconbitmap(str(ico))
    except Exception:
        pass

    gui = BackupGUI(root)
    root.geometry('920x640')
    root.mainloop()


if __name__ == '__main__':
    main()
