import os
import sys
import subprocess
import threading
import queue
import time
import atexit
import win32event
import win32api
import winerror
from datetime import datetime
from io import BytesIO

import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import scrolledtext

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
BOT_SCRIPT = "nobleprinter.py"
BOT_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────
# SINGLE INSTANCE CHECK (MUTEX)
# ─────────────────────────────────────────────
def check_single_instance():
    """Check if another instance is already running using Windows Mutex"""
    mutex_name = "NoblePrinterBot_Tray"
    try:
        mutex = win32event.CreateMutex(None, False, mutex_name)
        if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
            print("Another instance is already running. Exiting.")
            return False
        return True
    except Exception as e:
        print(f"Mutex check failed: {e}")
        return True

if not check_single_instance():
    sys.exit(0)

# ─────────────────────────────────────────────
# LOG VIEWER WINDOW
# ─────────────────────────────────────────────
class LogViewer:
    def __init__(self):
        self.root = None
        self.text_widget = None
        self.log_queue = queue.Queue()
        self.is_running = False
        self.window_visible = False
        
    def create_window(self):
        """Create the log viewer window (hidden by default)"""
        self.root = tk.Tk()
        self.root.title("Noble Printer Bot - Log Viewer")
        self.root.geometry("800x500")
        self.root.configure(bg='#2b2b2b')
        
        # Set window to not show on creation
        self.root.withdraw()
        
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Clear Log", command=self.clear_log)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Bot is running...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                              bd=1, relief=tk.SUNKEN, anchor=tk.W,
                              bg='#3c3c3c', fg='#ffffff')
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Toolbar
        toolbar = tk.Frame(self.root, bg='#3c3c3c', height=30)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        clear_btn = tk.Button(toolbar, text="Clear Log", command=self.clear_log,
                             bg='#4c4c4c', fg='white', padx=10)
        clear_btn.pack(side=tk.LEFT, padx=5, pady=2)
        
        # Text widget
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.text_widget = scrolledtext.ScrolledText(frame, 
                                                      wrap=tk.WORD,
                                                      bg='#1e1e1e',
                                                      fg='#d4d4d4',
                                                      insertbackground='white',
                                                      selectbackground='#264f78',
                                                      font=('Consolas', 10))
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Configure text tags for colors
        self.text_widget.tag_config('ERROR', foreground='#f48771')
        self.text_widget.tag_config('WARNING', foreground='#dcdcaa')
        self.text_widget.tag_config('INFO', foreground='#4ec9b0')
        
        self.is_running = True
        self.window_visible = False
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        # Start log reader
        self.update_logs()
        
    def update_logs(self):
        """Update log display with new entries"""
        if not self.is_running:
            return
            
        try:
            while True:
                log_entry = self.log_queue.get_nowait()
                self.insert_log(log_entry)
        except queue.Empty:
            pass
        
        if self.root:
            self.root.after(100, self.update_logs)
    
    def insert_log(self, log_entry):
        """Insert a log entry with appropriate color"""
        if not self.text_widget:
            return
            
        if 'ERROR' in log_entry or '❌' in log_entry:
            tag = 'ERROR'
        elif 'WARNING' in log_entry:
            tag = 'WARNING'
        else:
            tag = 'INFO'
        
        self.text_widget.insert(tk.END, log_entry + '\n', tag)
        self.text_widget.see(tk.END)
        
        if len(log_entry) > 80:
            self.status_var.set(log_entry[:77] + '...')
        else:
            self.status_var.set(log_entry)
    
    def clear_log(self):
        """Clear the log display"""
        if self.text_widget:
            self.text_widget.delete(1.0, tk.END)
    
    def hide_window(self):
        """Hide the window to system tray"""
        if self.root:
            self.root.withdraw()
            self.window_visible = False
    
    def show_window(self):
        """Show the window from system tray"""
        if self.root:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.window_visible = True
    
    def exit_app(self):
        """Exit the application"""
        self.is_running = False
        if self.root:
            self.root.quit()
            self.root.destroy()

# ─────────────────────────────────────────────
# BOT PROCESS MANAGER
# ─────────────────────────────────────────────
class BotProcessManager:
    def __init__(self, log_queue):
        self.log_queue = log_queue
        self.process = None
        self.is_running = False
        self.monitor_thread = None
        
    def start(self):
        """Start the bot as a subprocess"""
        if self.is_running:
            return False
            
        try:
            # Use pythonw.exe to run without console window
            pythonw_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            if not os.path.exists(pythonw_path):
                pythonw_path = sys.executable.replace("python.exe", "pythonw.exe")
            
            bot_path = os.path.join(BOT_DIR, BOT_SCRIPT)
            
            # Start subprocess with no window
            self.process = subprocess.Popen(
                [pythonw_path, bot_path],
                cwd=BOT_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=1,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW  # No console window
            )
            
            self.is_running = True
            self.monitor_thread = threading.Thread(target=self._monitor_output, daemon=True)
            self.monitor_thread.start()
            
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ Bot started (PID: {self.process.pid})")
            return True
            
        except Exception as e:
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ Failed to start bot: {e}")
            return False
    
    def _monitor_output(self):
        """Monitor subprocess output and send to log queue"""
        if not self.process:
            return
            
        try:
            for line in iter(self.process.stdout.readline, ''):
                if line:
                    self.log_queue.put(line.strip())
                if not self.is_running:
                    break
        except:
            pass
        
        if self.process.poll() is not None:
            self.is_running = False
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] ⚠️ Bot process ended (exit code: {self.process.returncode})")
            # Auto restart? You can enable this if desired
            # time.sleep(5)
            # self.start()
    
    def stop(self):
        """Stop the bot process"""
        if not self.is_running or not self.process:
            return
            
        try:
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] 🛑 Stopping bot...")
            self.process.terminate()
            
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait()
            
            self.is_running = False
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ Bot stopped")
            
        except Exception as e:
            self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ Error stopping bot: {e}")
    
    def restart(self):
        """Restart the bot"""
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] 🔄 Restarting bot...")
        self.stop()
        time.sleep(2)
        self.start()
    
    def status(self):
        """Get bot status"""
        if self.is_running and self.process and self.process.poll() is None:
            return "🟢 Running"
        else:
            return "🔴 Stopped"

# ─────────────────────────────────────────────
# TRAY ICON
# ─────────────────────────────────────────────
def create_icon():
    """Create a printer icon for system tray"""
    size = 64
    image = Image.new('RGB', (size, size), '#2b2b2b')
    draw = ImageDraw.Draw(image)
    
    draw.rectangle([10, 20, 54, 50], fill='#4c4c4c', outline='#ffffff')
    draw.rectangle([15, 25, 49, 45], fill='#2b2b2b')
    draw.rectangle([20, 35, 44, 42], fill='#4ec9b0')
    
    return image

# ─────────────────────────────────────────────
# MAIN TRAY APPLICATION
# ─────────────────────────────────────────────
class TrayApplication:
    def __init__(self):
        self.log_queue = queue.Queue()
        self.log_viewer = LogViewer()
        self.bot_manager = BotProcessManager(self.log_queue)
        self.icon = None
        
        self.log_viewer.log_queue = self.log_queue
        
    def run(self):
        """Run the tray application"""
        # Create log viewer (hidden by default)
        self.log_viewer.create_window()
        
        # Start bot automatically after a short delay
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Noble Printer Bot...")
        self.bot_manager.start()
        
        # Create system tray icon
        icon_image = create_icon()
        
        menu = pystray.Menu(
            pystray.MenuItem("📋 Show Log", lambda: self.log_viewer.show_window()),
            pystray.MenuItem("🔄 Restart Bot", lambda: self.bot_manager.restart()),
            pystray.MenuItem("▶️ Start Bot", lambda: self.bot_manager.start(), 
                            enabled=lambda item: not self.bot_manager.is_running),
            pystray.MenuItem("⏹️ Stop Bot", lambda: self.bot_manager.stop(), 
                            enabled=lambda item: self.bot_manager.is_running),
            pystray.MenuItem("📊 Status", lambda: self.show_status()),
            pystray.MenuItem("🚪 Exit", lambda: self.exit_app())
        )
        
        self.icon = pystray.Icon("noble_printer_bot", icon_image, "Noble Printer Bot", menu)
        
        # Run icon in separate thread
        icon_thread = threading.Thread(target=self.icon.run, daemon=True)
        icon_thread.start()
        
        # Start tkinter main loop
        self.log_viewer.root.mainloop()
    
    def show_status(self):
        """Show status in notification"""
        status = self.bot_manager.status()
        self.log_queue.put(f"[{datetime.now().strftime('%H:%M:%S')}] Status: {status}")
    
    def exit_app(self):
        """Exit the application"""
        self.bot_manager.stop()
        if self.icon:
            self.icon.stop()
        self.log_viewer.exit_app()

# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = TrayApplication()
    app.run()