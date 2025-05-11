import tkinter as tk
from tkinter import ttk, messagebox, Frame
from modules.sender import SenderModule
from modules.fetcher import FetcherModule
from modules.settings import SettingsModule, load_settings_env
import os
import threading

class ResponsiveApp:
    def __init__(self, root):
        self.root = root
        self.setup_ui()
        
    def setup_ui(self):
        # Configure main window
        self.root.title("Bulky-Mailman")
        self.root.geometry("900x700")
        
        # Add app icon if available
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass  # Icon not found, continue without it
            
        # Make the window content resize with the window
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        # Create main container frame
        main_frame = Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        
        # Create the notebook/tab control
        self.tab_control = ttk.Notebook(main_frame)
        self.tab_control.grid(row=0, column=0, sticky="nsew")
        
        # Create tabs with scrollable content
        self.tab_sender = self.add_scrollable_tab("📤 Bulk Sender")
        self.tab_fetcher = self.add_scrollable_tab("📥 Inbox Monitor")
        self.tab_settings = self.add_scrollable_tab("⚙️ Settings")
        
        # Initialize modules
        SenderModule(self.tab_sender)
        FetcherModule(self.tab_fetcher)
        SettingsModule(self.tab_settings)
        
        # Add status bar
        self.status_bar = ttk.Label(main_frame, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=1, column=0, sticky="ew")
        
        # Start background monitoring
        threading.Thread(target=FetcherModule.run_background_monitoring, daemon=True).start()
    
    def add_scrollable_tab(self, title):
        """Create a scrollable frame inside a tab"""
        # Create container for the tab
        tab = ttk.Frame(self.tab_control)
        self.tab_control.add(tab, text=title)
        
        # Configure container for expansion
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        
        # Create canvas for scrolling
        canvas = tk.Canvas(tab)
        canvas.grid(row=0, column=0, sticky="nsew")
        
        # Add vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Add horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tab, orient="horizontal", command=canvas.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        # Configure canvas
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # Create scrollable frame inside canvas
        scroll_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        return scroll_frame

def check_required_files():
    """Check for required files and directories"""
    # Create emails.csv if it doesn't exist
    if not os.path.exists("emails.csv"):
        try:
            import pandas as pd
            columns = ["From", "Subject", "Body", "Summary", "Suggested Response"]
            pd.DataFrame(columns=columns).to_csv("emails.csv", index=False)
        except Exception as e:
            print(f"Error creating emails.csv: {e}")

if __name__ == "__main__":
    # Load environment settings
    load_settings_env()
    
    # Create required files
    check_required_files()
    
    # Create root window
    root = tk.Tk()
    
    # Set theme for ttk - uncomment for a nicer look if available
    try:
        # Try to use a theme if themes are available
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
    except:
        pass  # Continue with default theme
    
    # Initialize app
    app = ResponsiveApp(root)
    
    # Show welcome message if email not configured
    if not os.getenv("EMAIL_WORK") or "example" in os.getenv("EMAIL_WORK"):
        messagebox.showinfo("Welcome", "⚙️ Please configure your email and AI settings in the Settings tab before using the tool.")
    
    # Start main loop
    root.mainloop()