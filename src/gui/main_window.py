import json
import logging
import shutil
import sys
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog
from tkinter import ttk, messagebox

import matplotlib
import matplotlib.dates as mdates  # Importing matplotlib.dates
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image, ImageTk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import os

# Import configuration settings from settings_manager.py
from config.settings_manager import (
    ASSETS_DIR,
    ENABLE_GOOGLE_SYNC,
    save_theme,
    default_config,
    PERSONAL_INFO_FILE,
    DATA_FILE_PATH,
    base_path,
    CONFIG_JSON_PATH,
    SERVICE_ACCOUNT_FILE
)
# Import utility functions for file I/O and Google Sheets synchronization
from src.utils.file_io import read_applications_from_excel, save_applications_to_excel
from src.utils.google_sheets import (
    read_from_google_sheets,
    write_to_google_sheets
)
# Import the centralized resource_path function from utils/utils.py
from src.utils.utils import resource_path

matplotlib.use('TkAgg')

def load_personal_info():
    """
    Loads personal information from a JSON file.
    Returns default data if the file does not exist.
    """
    if os.path.exists(PERSONAL_INFO_FILE):
        # Load JSON data if the file exists
        with open(PERSONAL_INFO_FILE, "r") as file:
            return json.load(file)
    else:
        # Return default information if file is missing
        return {
            "First Name": {"value": "John", "masked": False},
            "Last Name": {"value": "Doe", "masked": False},
            "Email": {"value": "john.doe@example.com", "masked": False},
            "Password": {"value": "password123", "masked": True},
            "Phone Number": {"value": "+1 (555) 123-4567", "masked": False},
            "Address Line 1": {"value": "123 Main St", "masked": False},
            "City": {"value": "Anytown", "masked": False},
            "State": {"value": "CA", "masked": False},
            "Zip Code": {"value": "12345", "masked": False},
            "Full Address": {"value": "123 Main St, Anytown, CA 12345", "masked": False},
            "University": {"value": "State University", "masked": False},
            "Degree": {"value": "BS in Computer Science", "masked": False},
        }


class ToolTip:
    """
    Creates a tooltip for a given widget.
    """

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Remove window decorations
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#FFFFE0", relief='solid', borderwidth=1,
                         font=("Arial", 10))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


class AppTrack(TkinterDnD.Tk):  # Inherit from TkinterDnD.Tk for drag-and-drop
    def __init__(self):
        super().__init__()

        # Initialize tooltip storage
        self.tooltips = []

        # --- 1. Initialize File Names and Paths ---
        self.mask_vars = {}
        self.APPLICATIONS_FILE_NAME = "Applications.xlsx"
        self.SERVICE_ACCOUNT_FILE_NAME = "service_account.json"
        self.app_base_path = os.path.dirname(os.path.abspath(__file__))
        self.CONFIG_DIR = os.path.join(self.app_base_path, 'config')
        self.DATA_DIR = os.path.join(self.app_base_path, 'Data')

        os.makedirs(self.CONFIG_DIR, exist_ok=True)
        os.makedirs(self.DATA_DIR, exist_ok=True)

        # Paths to config files
        self.CONFIG_JSON_PATH = os.path.join(self.CONFIG_DIR, 'app_config.json')
        self.PERSONAL_INFO_FILE = os.path.join(self.CONFIG_DIR, 'personal_info.json')

        # Initialize sync variable
        self.sync_to_google = False  # Default value; will be loaded from config

        # Initialize other variables
        self.status_combobox = None
        self.edit_entry = None
        self.menu_visible = False  # Variable to track menu visibility

        # --- 2. Standard Fonts ---
        self.label_font = ("Arial", 12)
        self.entry_font = ("Arial", 12)
        self.button_font = ("Arial", 12)
        self.heading_font = ("Arial", 14, "bold")

        # --- 3. Initialize Theme Variables ---
        self.theme = "Light"
        self.bg_color = "#F0F0F0"
        self.fg_color = "#000000"
        self.entry_bg_color = "#FFFFFF"
        self.entry_fg_color = "#000000"
        self.button_bg_color = "#E0E0E0"
        self.menu_bg_color = "#E0E0E0"
        self.menu_fg_color = "#000000"
        self.menu_active_bg = "#C0C0C0"

        # --- 4. Initialize Widget Tracking Lists ---
        self.entry_widgets = []
        self.clipboard_widgets = []  # Initialize to prevent AttributeError

        # --- 5. Configure the Main Window ---
        self.configure_window()

        # --- 6. Initialize Paths and Configuration ---
        self.initialize_paths()

        # --- 7. Load Clipboard Settings from Configuration ---
        self.load_clipboard_settings_from_config()

        # Add debug prints
        print(
            f"[DEBUG] After loading config: clipboard_enabled={self.clipboard_enabled}, clipboard_side={self.clipboard_side}")

        # --- 8. Create Custom Menu Bar ---
        self.create_custom_menu_bar()

        # --- 9. Load and Apply Theme ---
        self.load_and_apply_theme()

        # --- 10. Initialize Preferences ---
        self.initialize_preferences()

        # --- 11. Create UI Components ---
        self.create_ui_components()

        # --- 12. Load Assets (Images) ---
        self.load_assets()

        # --- 13. Initialize Additional GUI Components ---
        self.initialize_additional_gui()

        # --- 14. Load Application Data ---
        self.load_application_data()

        # --- 15. Setup the Main Layout ---
        self.setup_main_layout()

        # --- 16. Schedule Periodic Tasks ---
        self.schedule_tasks()

        # --- 17. Apply the Current Theme ---
        # Already handled by load_and_apply_theme

        # --- 18. Initialize and Schedule Sync Task ---
        self.sync_task = None
        self.schedule_sync()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        print("[DEBUG] AppTrack initialized successfully.")

    def draw_status_pie_chart_async(self, parent_frame):
        """Run the pie chart drawing in a separate thread."""
        threading.Thread(target=self.draw_status_pie_chart, args=(parent_frame,)).start()

    def configure_window(self):
        # Use the native title bar by removing overrideredirect
        self.title("AppTrack")
        try:
            icon_path = resource_path(os.path.join('assets', 'app_icon.png'))
            icon_image = tk.PhotoImage(file=icon_path)
            self.iconphoto(True, icon_image)  # Use .png file for the application icon
        except Exception as e:
            print(f"Error loading icon: {e}")
            logging.error(f"Error loading icon: {e}")

    def on_close(self):
        # Cancel any background tasks
        if self.sync_task:
            self.after_cancel(self.sync_task)
        # Close any threads or processes
        # Example: If using threading.Thread, call thread.join()
        self.destroy()
        sys.exit()

    def initialize_paths(self):
        self.BASE_PATH = base_path  # Use AppData base path directly
        self.CONFIG_DIR = os.path.join(self.BASE_PATH, 'config')
        self.CONFIG_JSON_PATH = CONFIG_JSON_PATH  # Already set to AppData path
        self.DATA_DIR = os.path.join(self.BASE_PATH, 'Data')
        self.DATA_FILE_PATH = DATA_FILE_PATH  # Directly use the AppData path for Applications.xlsx
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.CONFIG_DIR, exist_ok=True)
        print(f"[DEBUG] CONFIG_JSON_PATH set to: {self.CONFIG_JSON_PATH}")

    def load_and_apply_theme(self):
        """Load the theme from configuration and apply it."""
        try:
            if os.path.exists(self.CONFIG_JSON_PATH):
                with open(self.CONFIG_JSON_PATH, "r") as config_file:
                    config = json.load(config_file)
                    theme = config.get("theme", "Light")
                    print(f"[DEBUG] Loaded theme from config: {theme}")
            else:
                # If config file doesn't exist, set default theme and create config
                theme = "Light"
                self.update_config(theme=theme)
                print("[DEBUG] Config file not found. Set default theme to Light.")

            # Apply the loaded theme
            self.set_theme(theme)
        except json.JSONDecodeError:
            messagebox.showerror("Error", "app_config.json is corrupted. Reverting to default Light theme.")
            self.set_theme("Light")
            self.update_config(theme="Light")
        except Exception as e:
            print(f"Error loading theme: {e}")
            self.set_theme("Light")
            self.update_config(theme="Light")

    def initialize_preferences(self):
        """Initialize preferences like Google Sync based on the configuration."""
        self.sync_to_google = ENABLE_GOOGLE_SYNC
        self.google_sync_var = tk.BooleanVar(value=self.sync_to_google)

    def create_ui_components(self):
        self.create_custom_menu_bar()

    def load_assets(self):
        try:
            self.upload_xlsx_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_xlsx.png'))).resize((64, 64))
            )
            self.upload_json_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_json.png'))).resize((64, 64))
            )
            self.upload_sheets_id_icon = ImageTk.PhotoImage(
                Image.open(resource_path(os.path.join('assets', 'upload_sheets_id.png'))).resize((42, 42))
            )
            try:
                self.google_sync_icon = tk.PhotoImage(file=os.path.join(ASSETS_DIR, 'google_sync.png'))
            except Exception as e:
                print(f"Error loading google_sync.png: {e}")
                self.google_sync_icon = None

            try:
                self.applications_icon = tk.PhotoImage(file=os.path.join(ASSETS_DIR, 'applications.png'))
            except Exception as e:
                print(f"Error loading applications.png: {e}")
                self.applications_icon = None

        except Exception as e:
            print(f"Error loading assets: {e}")
            logging.error(f"Error loading assets: {e}")

    def initialize_additional_gui(self):
        self.selected_row = None
        self.selected_column = None
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.perform_search())
        self.applications_tree = None
        self.position_entry = None
        self.company_entry = None
        self.url_entry = None
        self.applications_df = pd.DataFrame()

    def load_application_data(self):
        """Loads application data from Applications.xlsx in AppData."""
        try:
            self.applications_df = read_applications_from_excel(self.DATA_FILE_PATH)
        except Exception as e:
            print(f"Error: Could not read the Excel file from AppData. {str(e)}")
            logging.error(f"Error: Could not read the Excel file from AppData. {str(e)}")
            self.applications_df = pd.DataFrame()

    def setup_main_layout(self):
        self.geometry("1250x700")  # Increased window size for better visibility
        self.main_paned_window = tk.PanedWindow(self, orient="horizontal")
        self.main_paned_window.pack(side='top', fill='both', expand=True)
        print("[DEBUG] Main PanedWindow created.")

        # Left Notebook with two tabs
        self.tab_control = ttk.Notebook(self.main_paned_window)
        self.add_application_tab = ttk.Frame(self.tab_control)
        self.view_edit_applications_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.add_application_tab, text="Add Application")
        self.tab_control.add(self.view_edit_applications_tab, text="View/Edit Applications")
        self.main_paned_window.add(self.tab_control, stretch="always")
        self.main_paned_window.paneconfigure(self.tab_control, minsize=1000)  # Set minimum size for main tabs
        print("[DEBUG] Main tab control with 'Add Application' and 'View/Edit Applications' tabs added.")

        # Clipboard Notebook
        self.clipboard_notebook = ttk.Notebook(self.main_paned_window)
        self.clipboard_tab = ttk.Frame(self.clipboard_notebook)
        self.clipboard_notebook.add(self.clipboard_tab, text="Clipboard")
        print("[DEBUG] Clipboard notebook created with 'Clipboard' tab.")

        # Debug info
        print(
            f"[DEBUG] Configuring main layout. Clipboard Enabled: {self.clipboard_enabled}, Clipboard Side: {self.clipboard_side}")

        # Conditionally add Clipboard Notebook based on clipboard_enabled
        if self.clipboard_enabled:
            if self.clipboard_side == 'right':
                self.main_paned_window.add(self.clipboard_notebook, stretch="always")
                self.main_paned_window.paneconfigure(self.clipboard_notebook, minsize=250, width=250)  # Enforce width
                print("[DEBUG] Clipboard notebook added to the right side with fixed width.")
            elif self.clipboard_side == 'left':
                self.main_paned_window.add(self.clipboard_notebook, stretch="always", before=self.tab_control)
                self.main_paned_window.paneconfigure(self.clipboard_notebook, minsize=250, width=250)  # Enforce width
                print("[DEBUG] Clipboard notebook added to the left side with fixed width.")
        else:
            print("[DEBUG] Clipboard feature is disabled; Clipboard tab not added.")

        # Create UI components
        self.create_add_application_tab()
        self.create_view_edit_applications_tab()
        self.create_personal_info_tab()

        # Update theme for all widgets within the tabs
        self.update_all_widgets_theme(self.add_application_tab)
        self.update_all_widgets_theme(self.view_edit_applications_tab)
        if self.clipboard_enabled:
            self.update_all_widgets_theme(self.clipboard_tab)

        print("[DEBUG] Main layout setup completed.")

    def on_add_application_tab_resize(self, event):
        """
        Adjusts the pie chart Canvas size based on the parent frame's size.
        """
        try:
            canvas_width = min(event.width, 500)  # Limit maximum width
            canvas_height = min(event.height, 300)  # Limit maximum height
            if hasattr(self, 'pie_canvas') and self.pie_canvas.winfo_exists():
                self.pie_canvas.config(width=canvas_width, height=canvas_height)
                self.draw_status_pie_chart()  # Redraw to fit new size
                print("[DEBUG] Pie chart Canvas resized and redrawn.")
        except Exception as e:
            print(f"Error resizing pie chart Canvas: {e}")
            logging.error(f"Error resizing pie chart Canvas: {e}")

    def schedule_tasks(self):
        if self.sync_to_google:
            self.after(60000, self.sync_to_google_sheets)
            self.schedule_sync()
            self.apply_theme()

    def load_clipboard_settings_from_config(self):
        """Load clipboard settings from the configuration."""
        try:
            if os.path.exists(self.CONFIG_JSON_PATH):
                with open(self.CONFIG_JSON_PATH, "r") as config_file:
                    config = json.load(config_file)
                    self.clipboard_side = config.get("clipboard_side", "right")
                    self.clipboard_enabled = config.get("clipboard_enabled", True)
                    self.clipboard_width = config.get("clipboard_width", 250)  # Default width to 250
                    print(f"[DEBUG] Loaded configuration from {self.CONFIG_JSON_PATH}:")
                    print(f"[DEBUG] clipboard_enabled = {self.clipboard_enabled}")
                    print(f"[DEBUG] clipboard_side = {self.clipboard_side}")
            else:
                # If config file doesn't exist, set defaults and create config
                self.clipboard_side = "right"
                self.clipboard_enabled = True
                self.update_config(clipboard_side=self.clipboard_side, clipboard_enabled=self.clipboard_enabled)
                print("[DEBUG] Config file not found. Set clipboard_side to 'right' and clipboard_enabled to True.")
        except json.JSONDecodeError:
            # Backup the corrupted config file
            backup_path = self.CONFIG_JSON_PATH + ".bak"
            shutil.copy(self.CONFIG_JSON_PATH, backup_path)
            print(f"[DEBUG] Corrupted config file backed up to {backup_path}")
            messagebox.showerror("Error", "app_config.json is corrupted. Reverting to default settings.")
            self.clipboard_side = "right"
            self.clipboard_enabled = True
            self.update_config(clipboard_side=self.clipboard_side, clipboard_enabled=self.clipboard_enabled)
            print("[DEBUG] app_config.json is corrupted. Reverted to default clipboard settings.")
        except Exception as e:
            print(f"Error loading clipboard settings: {e}")
            self.clipboard_side = "right"
            self.clipboard_enabled = True
            self.update_config(clipboard_side=self.clipboard_side, clipboard_enabled=self.clipboard_enabled)
            print("[DEBUG] Error loading clipboard settings. Reverted to default settings.")

    def load_theme_from_config(self):
        """Load the saved theme setting from app_config.json in AppData, or create it with default settings if it doesn't exist."""
        # Check if app_config.json exists in AppData (CONFIG_JSON_PATH)
        if not os.path.exists(CONFIG_JSON_PATH):
            # Create app_config.json in AppData with default theme settings
            default_config = {"theme": "Light"}
            with open(CONFIG_JSON_PATH, "w") as config_file:
                json.dump(default_config, config_file, indent=4)
            return False  # Default to Light theme

        # Load theme from existing app_config.json in AppData
        try:
            with open(CONFIG_JSON_PATH, "r") as config_file:
                config = json.load(config_file)
                return config.get("theme", "Light") == "Dark"
        except json.JSONDecodeError:
            messagebox.showerror("Error", "app_config.json is corrupted. Reverting to default settings.")
            return False  # Default to Light theme

    def create_entry(self, parent, **kwargs):
        """
        Helper function to create a tk.Entry widget with theme-aware cursor and background colors.
        """
        # Remove unsupported options for tk.Entry
        unsupported_options = ['wraplength']
        for option in unsupported_options:
            if option in kwargs:
                print(f"[WARNING] Option '{option}' is not supported by tk.Entry and will be ignored.")
                logging.warning(f"Option '{option}' is not supported by tk.Entry and will be ignored.")
                kwargs.pop(option)

        # Automatically set bg and fg based on the current theme
        kwargs.setdefault('bg', self.entry_bg_color)
        kwargs.setdefault('fg', self.entry_fg_color)
        kwargs.setdefault('insertbackground', self.fg_color)  # Cursor color

        entry = tk.Entry(parent, **kwargs)
        self.entry_widgets.append(entry)  # Keep track for future updates
        return entry

    def create_add_application_tab(self):
        """
        Enhances the 'Add Application' tab by creating a master panel that contains
        labels and entry fields for Company, Position, and Application Portal URL,
        along with a Submit button. The fields are arranged side by side with labels on top.
        Also includes a pie chart and a line graph side by side below the input fields.
        """
        # --- Step 1: Configure Parent Grid ---
        self.add_application_tab.columnconfigure(0, weight=1)  # Left padding
        self.add_application_tab.columnconfigure(1, weight=0)  # Master panel
        self.add_application_tab.columnconfigure(2, weight=1)  # Right padding

        # Configure rows: input fields and graphs
        self.add_application_tab.rowconfigure(0, weight=0)  # Input fields row
        self.add_application_tab.rowconfigure(1, weight=1)  # Graphs row

        # --- Step 2: Create the Master Frame ---
        master_frame = tk.Frame(
            self.add_application_tab,
            bg=self.bg_color,
            padx=20,  # Horizontal padding inside the master frame
            pady=20  # Vertical padding inside the master frame
        )
        master_frame.grid(row=0, column=1, sticky="n")  # Aligned to the top, centered horizontally

        # --- Step 3: Configure Master Frame Grid ---
        master_frame.columnconfigure(0, weight=1, uniform="col")  # Column 0: First field
        master_frame.columnconfigure(1, weight=1, uniform="col")  # Column 1: Second field
        master_frame.columnconfigure(2, weight=1, uniform="col")  # Column 2: Third field

        # Define a larger font size for labels
        label_font = (self.label_font[0], 14)  # Increase font size to 14

        # --- Create Fields Side by Side ---

        # --- Column 0: Company Label and Entry ---
        company_label = tk.Label(
            master_frame,
            text="Company:",
            font=label_font,
            bg=self.bg_color,
            fg=self.fg_color,
            anchor="w"
        )
        company_label.grid(row=0, column=0, sticky="w", padx=(10, 10))  # Label on top
        self.company_entry = self.create_entry(
            master_frame,
            font=self.entry_font,
            width=30,
            bg=self.entry_bg_color,
            fg=self.entry_fg_color
        )
        self.company_entry.grid(row=1, column=0, sticky="ew", padx=(10, 10))  # Entry below label

        # --- Column 1: Position Label and Entry ---
        position_label = tk.Label(
            master_frame,
            text="Position:",
            font=label_font,
            bg=self.bg_color,
            fg=self.fg_color,
            anchor="w"
        )
        position_label.grid(row=0, column=1, sticky="w", padx=(10, 10))  # Label on top
        self.position_entry = self.create_entry(
            master_frame,
            font=self.entry_font,
            width=30,
            bg=self.entry_bg_color,
            fg=self.entry_fg_color
        )
        self.position_entry.grid(row=1, column=1, sticky="ew", padx=(10, 10))  # Entry below label

        # --- Column 2: URL Label and Entry ---
        url_label = tk.Label(
            master_frame,
            text="Application Portal URL:",
            font=label_font,
            bg=self.bg_color,
            fg=self.fg_color,
            anchor="w"
        )
        url_label.grid(row=0, column=2, sticky="w", padx=(10, 10))  # Label on top
        self.url_entry = self.create_entry(
            master_frame,
            font=self.entry_font,
            width=30,
            bg=self.entry_bg_color,
            fg=self.entry_fg_color
        )
        self.url_entry.grid(row=1, column=2, sticky="ew", padx=(10, 10))  # Entry below label

        # --- Submit Button Centered Below Fields ---
        submit_button = ttk.Button(
            master_frame,
            text="Submit",
            command=self.save_application,
            style="Custom.TButton",
            width=6  # Adjust width as needed
        )
        submit_button.grid(row=2, column=0, columnspan=3, pady=(20, 0))  # Spanning all columns

        # --- Step 4: Create a Container Frame for Graphs ---
        graphs_frame = tk.Frame(
            self.add_application_tab,
            bg=self.bg_color,
            padx=10,  # Padding around the graphs
            pady=10
        )
        graphs_frame.grid(row=1, column=1, sticky="n")  # Positioned below the master_frame

        # Configure graphs_frame grid to have two columns: one for pie chart, one for line graph
        graphs_frame.columnconfigure(0, weight=1)
        graphs_frame.columnconfigure(1, weight=1)
        graphs_frame.rowconfigure(0, weight=1)

        # --- Step 5: Create Frames for Pie Chart and Line Graph within graphs_frame ---
        # Frame for Pie Chart
        self.pie_chart_frame = tk.Frame(
            graphs_frame,
            bg=self.bg_color,
            padx=50,
            pady=5
        )
        self.pie_chart_frame.grid(row=0, column=1, sticky="nsew")  # Left side

        # Frame for Line Graph
        self.line_graph_frame = tk.Frame(
            graphs_frame,
            bg=self.bg_color,
            padx=50,
            pady=5
        )
        self.line_graph_frame.grid(row=0, column=0, sticky="nsew")  # Right side

        # --- Step 6: Draw the Pie Chart and Line Graph Inside Their Respective Frames ---
        self.draw_status_pie_chart(self.pie_chart_frame)
        self.draw_submissions_line_graph(self.line_graph_frame)

    def create_view_edit_applications_tab(self):
        self.view_edit_applications_tab.columnconfigure(0, weight=1)
        self.view_edit_applications_tab.rowconfigure(1, weight=1)

        search_frame = tk.Frame(self.view_edit_applications_tab, bg=self.bg_color)
        search_frame.grid(row=0, column=0, sticky="ew", padx=7, pady=7)

        # Search label and entry field
        ttk.Label(search_frame, text="Search:", font=self.label_font).pack(side="left", padx=(15, 5))
        search_entry = self.create_entry(search_frame, textvariable=self.search_var, width=30, font=self.entry_font)
        search_entry.pack(side="left", padx=(5, 5))

        # Frame for the main Treeview
        treeview_frame = tk.Frame(self.view_edit_applications_tab, bg=self.bg_color)
        treeview_frame.grid(row=1, column=0, sticky="nsew", padx=2, pady=(0, 2))
        self.view_edit_applications_tab.rowconfigure(1, weight=1)
        self.view_edit_applications_tab.columnconfigure(0, weight=1)

        # Define Treeview columns
        columns = ("No", "Company", "Position", "Application Portal URL", "Date Applied", "Status")
        self.applications_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", selectmode="extended",
                                              height=15)

        # Configure each column explicitly
        for col in columns:
            self.applications_tree.heading(
                col,
                text=col,
                anchor="center",
                command=lambda _col=col: self.sort_treeview_column(_col, False)
            )
            self.applications_tree.column(
                col,
                anchor="w" if col not in ["No", "Date Applied", "Status"] else "center",
                stretch=True,
                width=5 if col == "No" else 175 if col == "Company" else 120 if col == "Position" else 200 if col == "Application Portal URL" else 80 if col == "Date Applied" else 70 if col == "Status" else 80
            )

        # Apply a consistent font to Treeview rows
        style = ttk.Style()
        style.configure("Treeview", font=self.entry_font)
        style.configure("Treeview.Heading", font=self.heading_font)

        # Bind Treeview events for clicking and context menu based on OS
        if sys.platform.startswith('darwin'):
            # macOS bindings
            self.applications_tree.bind("<Button-2>", self.show_context_menu)  # Secondary click (two-finger click)
            self.applications_tree.bind("<Control-Button-1>", self.show_context_menu)  # Control-click
        else:
            # Windows and other OS bindings
            self.applications_tree.bind("<Button-3>", self.show_context_menu)  # Right-click

        # Bind left-click for selection
        self.applications_tree.bind("<Button-1>", self.on_treeview_click)

        # Bind keyboard shortcut for context menu
        self.applications_tree.bind("<Shift-F10>", self.show_context_menu)

        # Bind double-click event to open URLs
        self.applications_tree.bind("<Double-1>", self.on_treeview_double_click)  # <--- INSERTED HERE

        # Add vertical scrollbar for Treeview
        vsb = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.applications_tree.yview)
        self.applications_tree.configure(yscrollcommand=vsb.set)

        # Position Treeview and scrollbar in grid
        self.applications_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Configure grid weights for scrollbar frame
        treeview_frame.rowconfigure(0, weight=1)
        treeview_frame.columnconfigure(0, weight=1)

        # Populate the Treeview with data
        self.populate_treeview(self.applications_df)  # Ensure binding is active before populating data

    def create_personal_info_tab(self):
        """
        Redesign the 'Personal Information' tab with fixed-sized content boxes,
        minimal spacing between main labels and info labels, and an enlarged 📋 emoji
        that users can click to copy information. Masking is handled via the clipboard editor.
        """

        # Clear existing content in the clipboard_tab
        for widget in self.clipboard_tab.winfo_children():
            widget.destroy()

        # Reset clipboard_widgets to prevent accumulation
        self.clipboard_widgets = []

        # Create a Canvas without a scrollbar
        self.clipboard_canvas = tk.Canvas(self.clipboard_tab, bg=self.bg_color, highlightthickness=0)
        self.clipboard_canvas.pack(side="top", fill="both", expand=True, padx=10, pady=10, anchor="center")

        # Create a frame inside the canvas to hold all content
        content_frame = tk.Frame(self.clipboard_canvas, bg=self.bg_color, highlightthickness=0)
        self.clipboard_canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Configure grid to expand properly
        content_frame.columnconfigure(0, weight=1)

        personal_info = load_personal_info()

        # Store entries for saving later
        self.personal_info_entries = {}

        # Store references to all clipboard widgets for theme updates
        self.clipboard_widgets = []  # Reset the list

        # Adjustable settings
        settings = {
            # Padding and Spacing
            "content_padding": {"padx": 5, "pady": 1},
            "frame_spacing": {"padx": 5, "pady": 3},  # Reduced from 7 to 5
            "text_frame_padding": {"padx": 0, "pady": 0},  # Minimal padding for text_frame
            "emoji_frame_padding": {"padx": 5, "pady": 5},  # Padding around the emoji

            "label_info_spacing": 0,  # Minimal spacing between main label and info label

            # Fixed Dimensions for Info Frames
            "info_frame_width": 215,  # Fixed width in pixels
            "info_frame_height": 45,  # Increased to accommodate larger emoji

            # Fonts
            "main_label_font": ("TkDefaultFont", 12, "bold"),
            "info_label_font": ("TkDefaultFont", 9, "italic"),  # Set to italic
            "copy_emoji_font": ("TkDefaultFont", 18),  # Increased for larger 📋

            # Styles
            "frame_border": {"relief": "flat", "bd": 0, "highlightthickness": 0},  # Removed borders

            # Label Configuration
            "wrap_length": 250,

            # Colors
            "bg_color": self.bg_color,
            "fg_color": self.fg_color,
            "entry_bg_color": self.entry_bg_color,
        }

        # Initialize row counter for grid placement
        current_row = 0

        # Iterate over personal_info to create labels, entry fields, and masking checkboxes
        for idx, (label, info_dict) in enumerate(personal_info.items()):
            value = info_dict.get("value", "")
            masked = info_dict.get("masked", False)

            # Determine display value based on masking
            display_value = self.mask_value(value) if masked else value

            # Frame for each personal info block with fixed size and no borders
            info_frame = tk.Frame(
                content_frame,
                bg=self.entry_bg_color,  # Use entry_bg_color for distinct box appearance
                width=settings["info_frame_width"],
                height=settings["info_frame_height"],
                **settings["frame_border"],  # Updated to have no border
            )
            info_frame.grid(row=current_row, column=0, sticky="ew", **settings["frame_spacing"])
            info_frame.grid_propagate(False)  # Prevent frame from resizing to fit content
            info_frame.columnconfigure(0, weight=0)  # Button column doesn't expand
            info_frame.columnconfigure(1, weight=1)  # Text column expands

            # Store reference for theme updates
            self.clipboard_widgets.append(info_frame)

            # Create a horizontal layout: text_frame on the left, emoji_frame on the right
            text_frame = tk.Frame(info_frame, bg=self.entry_bg_color, **settings["text_frame_padding"])
            text_frame.grid(row=0, column=1, sticky="w", padx=(0, 5))  # Allow text to expand
            text_frame.columnconfigure(0, weight=1)  # Ensure the text expands properly

            emoji_frame = tk.Frame(info_frame, bg=self.entry_bg_color, **settings["emoji_frame_padding"])
            emoji_frame.grid(row=0, column=0, sticky="w", padx=(5, 0))  # Left-align with padding
            emoji_frame.columnconfigure(0, weight=1)
            emoji_frame.rowconfigure(0, weight=1)

            # Store references
            self.clipboard_widgets.extend([text_frame, emoji_frame])

            # Main label
            label_widget = ttk.Label(
                text_frame,
                text=label,
                font=settings["main_label_font"],
                anchor="w",
            )
            label_widget.pack(anchor="w")
            self.clipboard_widgets.append(label_widget)

            # Info label directly underneath the main label
            info_label = ttk.Label(
                text_frame,
                text=display_value,
                font=settings["info_label_font"],  # Italic font
                anchor="w",
                wraplength=settings["wrap_length"],
                justify="left",
            )
            info_label.pack(anchor="w", pady=(settings["label_info_spacing"], 0))
            self.personal_info_entries[label] = info_label
            self.clipboard_widgets.append(info_label)

            # 📋 Emoji as a clickable label, vertically centered
            copy_emoji = tk.Label(  # Changed from ttk.Label to tk.Label for more control
                emoji_frame,
                text="📄",
                font=settings["copy_emoji_font"],
                bg=self.entry_bg_color,  # Match the background to blend seamlessly
                fg=self.fg_color,
                cursor="hand2",  # Change cursor to indicate interactivity
            )
            copy_emoji.pack(expand=True, fill='both')
            self.clipboard_widgets.append(copy_emoji)

            # Bind click event to the 📋 emoji
            copy_emoji.bind(
                "<Button-1>",
                lambda e, lbl=label: self.copy_to_clipboard(self.get_actual_personal_info_value(lbl))
            )

            # Optional: Bind hover events to change emoji color on hover
            copy_emoji.bind("<Enter>", lambda e, ce=copy_emoji: ce.config(fg="#5b5b5b"))  # Example: red on hover
            copy_emoji.bind("<Leave>", lambda e, ce=copy_emoji: ce.config(fg=self.fg_color))  # Revert on leave

            # Increment row counter
            current_row += 1

        # Ensure all rows expand equally
        content_frame.update_idletasks()  # Ensure all widgets are rendered
        self.clipboard_canvas.config(scrollregion=self.clipboard_canvas.bbox("all"))

        for i in range(current_row + 1):
            content_frame.rowconfigure(i, weight=1)

        # Bind the <Configure> event to update the scrollregion
        # Since we've removed the scrollbar, this line can be kept or removed based on preference
        content_frame.bind("<Configure>", lambda e: self.clipboard_canvas.config(scrollregion=self.clipboard_canvas.bbox("all")))

    def create_custom_menu_bar(self):
        """Create a custom menu bar with theme-aware styling."""
        print("[DEBUG] Creating custom menu bar...")

        # Destroy existing menu bar if it exists
        if hasattr(self, 'menu_bar') and self.menu_bar:
            self.menu_bar.destroy()
            print("[DEBUG] Existing menu bar destroyed.")

        # Main menu bar frame with no border or highlight
        self.menu_bar = tk.Frame(self, bg=self.menu_bg_color, height=25, bd=0, highlightthickness=0)
        self.menu_bar.pack(side='top', fill='x')
        print("[DEBUG] New menu bar frame created.")

        # Define styles for buttons
        style = ttk.Style()
        style.configure("Settings.TButton",
                        background=self.menu_bg_color,
                        foreground=self.menu_fg_color,
                        font=("Arial", 12),
                        relief="flat")
        print("[DEBUG] Styles for Settings.TButton configured.")

        # Dropdown menu for Settings
        self.settings_menu = tk.Menu(
            self.menu_bar,
            tearoff=0,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        print("[DEBUG] Settings dropdown menu created.")

        # Add main Settings commands
        self.settings_menu.add_command(label="Applications File", command=self.open_applications_config_dialog)
        self.settings_menu.add_command(label="Google Sync Settings", command=self.open_settings_dialog)
        print("[DEBUG] Main Settings menu commands added.")

        # Add a separator before Theme and Clipboard submenus
        self.settings_menu.add_separator()
        print("[DEBUG] Separator added to Settings menu.")

        # --- Theme Submenu ---
        self.theme_menu = tk.Menu(
            self.settings_menu,
            tearoff=0,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        print("[DEBUG] Theme submenu created.")

        # Add Light Mode and Dark Mode options to Theme submenu
        self.theme_menu.add_command(label="Light Mode", command=lambda: self.set_theme("Light"))
        self.theme_menu.add_command(label="Dark Mode", command=lambda: self.set_theme("Dark"))
        print("[DEBUG] 'Light Mode' and 'Dark Mode' commands added to Theme submenu.")

        # Add Theme submenu to Settings menu
        self.settings_menu.add_cascade(label="Theme", menu=self.theme_menu)
        print("[DEBUG] Theme submenu cascaded under Settings menu.")

        # --- Clipboard Submenu ---
        self.clipboard_menu = tk.Menu(
            self.settings_menu,
            tearoff=0,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        print("[DEBUG] Clipboard submenu created.")

        # Add separator for better organization
        self.clipboard_menu.add_separator()

        # Initialize the clipboard_enabled_var before using it
        self.clipboard_enabled_var = tk.BooleanVar(value=self.clipboard_enabled)

        # Add Clipboard Checkbutton
        self.clipboard_menu.add_checkbutton(
            label="Clipboard",
            variable=self.clipboard_enabled_var,
            command=self.toggle_clipboard_feature,
        )
        print("[DEBUG] 'Clipboard' Checkbutton added to Clipboard submenu.")

        # Add "Edit Clipboard" command
        self.clipboard_menu.add_command(
            label="Edit Clipboard",
            command=self.open_clipboard_editor,
        )
        print("[DEBUG] 'Edit Clipboard' command added to Clipboard submenu.")

        # Create Clipboard Position submenu
        self.clipboard_position_menu = tk.Menu(
            self.clipboard_menu,
            tearoff=0,
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color
        )
        print("[DEBUG] Clipboard Position submenu created under Clipboard submenu.")

        # Add "Left" and "Right" commands to Clipboard Position submenu
        self.clipboard_position_menu.add_command(
            label="Left",
            command=lambda: self.move_clipboard_tab('left'),
        )
        self.clipboard_position_menu.add_command(
            label="Right",
            command=lambda: self.move_clipboard_tab('right'),
        )
        print("[DEBUG] 'Left' and 'Right' commands added to Clipboard Position submenu.")

        # Add Clipboard Position submenu to Clipboard submenu
        self.clipboard_menu.add_cascade(
            label="Clipboard Position",
            menu=self.clipboard_position_menu
        )
        print("[DEBUG] Clipboard Position submenu cascaded under Clipboard submenu.")

        # Add Clipboard submenu to Settings menu
        self.settings_menu.add_cascade(label="Clipboard", menu=self.clipboard_menu)
        print("[DEBUG] Clipboard submenu cascaded under Settings menu.")

        # Settings button with gear emoji
        self.settings_button = ttk.Button(
            self.menu_bar,
            text='⚙️',  # Gear emoji as settings icon
            style="Settings.TButton",
            width=3,
            command=lambda: self.settings_menu.post(self.winfo_pointerx(), self.winfo_pointery())
        )
        self.settings_button.pack(side='left', padx=5)
        print("[DEBUG] Settings button with gear emoji added to menu bar.")

        # Styled Google Sync Toggle Checkbutton using tk.Checkbutton
        self.google_sync_var = tk.BooleanVar(value=self.sync_to_google)
        self.google_sync_checkbutton = tk.Checkbutton(
            self.menu_bar,
            text="Google Sync",
            variable=self.google_sync_var,
            command=self.toggle_sync,
            activebackground=self.menu_active_bg,  # Background when active/hovered
            activeforeground=self.menu_fg_color,  # Text color when active/hovered
            selectcolor=self.menu_bg_color,  # Color of the indicator when selected
            borderwidth=0,  # Remove border
            highlightthickness=0,  # Remove highlight
            relief="flat",  # Flat relief to avoid 3D effects
            padx=2,  # Minimal horizontal padding
            pady=0,  # No vertical padding
            bg=self.menu_bg_color,  # Match the menu bar background
            fg=self.menu_fg_color  # Match the text color
        )
        self.google_sync_checkbutton.pack(side='left', padx=(0, 5), pady=0, anchor='center')
        print("[DEBUG] Google Sync Checkbutton added to menu bar.")

        print("[DEBUG] Custom menu bar created successfully.")

    def move_clipboard_tab(self, side):
        """
        Move the clipboard tab to the specified side ('left' or 'right').
        If the clipboard is currently disabled, re-enable it and update the settings menu.
        """
        if side not in ['left', 'right']:
            messagebox.showerror("Invalid Option",
                                 "Please select either 'left' or 'right' for the clipboard tab position.")
            return

        # If clipboard is disabled, enable it
        if not self.clipboard_enabled:
            self.clipboard_enabled_var.set(True)  # Check the clipboard checkbox in the settings menu
            self.toggle_clipboard_feature()  # Enable the clipboard feature

        # Remove the clipboard_notebook from the PanedWindow
        self.main_paned_window.forget(self.clipboard_notebook)

        # Re-add the clipboard_notebook to the specified side
        if side == 'right':
            self.main_paned_window.add(self.clipboard_notebook, stretch="always")
        elif side == 'left':
            self.main_paned_window.add(self.clipboard_notebook, stretch="always", before=self.tab_control)

        # Enforce fixed width for the clipboard tab
        self.clipboard_notebook.update_idletasks()
        self.clipboard_notebook.config(width=250)
        self.main_paned_window.paneconfigure(self.clipboard_notebook, minsize=250, width=250)

        # Update the clipboard_side attribute and save it to the config
        self.clipboard_side = side
        self.update_config(clipboard_side=side)

        # Refresh the layout
        self.main_paned_window.update_idletasks()
        print(f"[DEBUG] Moved Clipboard tab to the {side} side with a fixed width of 250.")

    def toggle_clipboard_feature(self):
        """
        Toggle the Clipboard feature on or off based on the checkbox state.
        """
        state = self.clipboard_enabled_var.get()
        print(f"[DEBUG] Clipboard feature toggled to {'enabled' if state else 'disabled'}.")

        # Update the configuration
        self.update_config(clipboard_enabled=state)

        if state:
            # Enable Clipboard: Show the Clipboard tab
            self.show_clipboard_tab()
        else:
            # Disable Clipboard: Hide the Clipboard tab
            self.hide_clipboard_tab()

    def show_clipboard_tab(self):
        """Show the Clipboard tab in the main layout with strict width control."""
        print("[DEBUG] Showing Clipboard tab.")

        # Ensure the Clipboard notebook is added to the PanedWindow
        if not self.clipboard_notebook.winfo_ismapped():
            if self.clipboard_side == 'right':
                # Add clipboard on the right side
                self.main_paned_window.add(self.clipboard_notebook, stretch="always")
            elif self.clipboard_side == 'left':
                # Add clipboard on the left side
                self.main_paned_window.add(self.clipboard_notebook, stretch="always", before=self.tab_control)

            # Enforce fixed size for the clipboard tab
            self.clipboard_notebook.update_idletasks()
            self.clipboard_notebook.config(width=250)
            self.main_paned_window.paneconfigure(self.clipboard_notebook, minsize=250, width=250)

            # Ensure the main notebook always has a minimum width
            self.tab_control.update_idletasks()
            self.main_paned_window.paneconfigure(self.tab_control, minsize=1000, weight=1)

            # Update the layout
            self.main_paned_window.update_idletasks()
            print("[DEBUG] Clipboard tab shown with a fixed width of 250.")

    def hide_clipboard_tab(self):
        """Hide the Clipboard tab from the main layout."""
        print("[DEBUG] Hiding Clipboard tab.")
        self.main_paned_window.forget(self.clipboard_notebook)
        self.main_paned_window.update_idletasks()
        print("[DEBUG] Clipboard tab removed from the main layout.")

    def draw_status_pie_chart(self, parent_frame):
        """
        Draws a pie chart representing the distribution of application statuses
        directly on a Tkinter Canvas within the specified parent frame.

        Parameters:
        - parent_frame (tk.Frame): The frame where the pie chart Canvas will be placed.
        """
        try:
            # Check if 'Status' column exists
            if 'Status' not in self.applications_df.columns:
                messagebox.showwarning("Missing Data",
                                       "'Status' column is missing from the data. Pie chart cannot be rendered.")
                return

            # Check if 'Status' column has non-zero entries
            if self.applications_df['Status'].dropna().empty:
                return

            # Count the number of applications in each status
            status_counts = self.applications_df['Status'].value_counts()

            # Define all possible statuses to ensure consistency in the pie chart
            all_statuses = ['Submitted', 'Rejected', 'Interview', 'Offer']
            sizes = [status_counts.get(status, 0) for status in all_statuses]

            # Calculate total to determine percentages
            total = sum(sizes)
            if total == 0:
                messagebox.showwarning("No Data", "There are no applications to display in the pie chart.")
                return  # Skip drawing the pie chart if there's no data

            # Remove any existing pie chart Canvas to avoid overlap
            if hasattr(self, 'pie_canvas') and self.pie_canvas.winfo_exists():
                self.pie_canvas.destroy()
                self.pie_canvas = None  # Clear the reference

            # Create a new Canvas for the pie chart with increased size
            self.pie_canvas = tk.Canvas(
                parent_frame,
                width=350,  # Increased width by 20%
                height=200,  # Increased height to accommodate total label
                bg=self.bg_color,  # Use the window's background color
                highlightthickness=0
            )
            self.pie_canvas.grid(row=0, column=0, padx=5, pady=5)  # Place using grid with padding

            # Define colors for each status
            colors = ['#9fc5e8', '#ea9999', '#b6d7a8', '#ffe599']  # Submitted, Rejected, Interview, Offer

            # Starting angle
            start_angle = 0

            # Draw pie slices
            for i, (status, size) in enumerate(zip(all_statuses, sizes)):
                if size == 0:
                    continue  # Skip if no applications in this status

                # Calculate the angle for this slice
                extent = (size / total) * 360

                # Draw the slice with increased bounding box
                self.pie_canvas.create_arc(
                    10, 10, 174, 174,  # Increased bounding box (145 * 1.2 ≈ 174)
                    start=start_angle,
                    extent=extent,
                    fill=colors[i],
                    outline=self.bg_color
                )

                # Update the starting angle for the next slice
                start_angle += extent

            # Draw the legend
            legend_start_x = 200  # Adjusted for increased canvas
            legend_start_y = 10
            legend_spacing = 20  # Increased spacing for better readability

            for i, status in enumerate(all_statuses):
                size = sizes[i]
                if size == 0:
                    continue  # Skip statuses with zero count

                # Draw color box
                self.pie_canvas.create_rectangle(
                    legend_start_x, legend_start_y + i * legend_spacing,
                                    legend_start_x + 15, legend_start_y + i * legend_spacing + 15,  # Increased size
                    fill=colors[i],
                    outline=colors[i]
                )

                # Draw status label with theme-aware foreground color
                label_text = f"{status} ({size})"
                self.pie_canvas.create_text(
                    legend_start_x + 20, legend_start_y + i * legend_spacing + 7.5,  # Centered vertically
                    text=label_text,
                    anchor='w',
                    fill=self.fg_color,  # Use theme-dependent foreground color
                    font=("Arial", 10)  # Slightly increased font size
                )

            # Add the "Total Applications" label to the legend
            self.pie_canvas.create_text(
                legend_start_x, legend_start_y + len(all_statuses) * legend_spacing + 10,
                text=f"Total Applications: ({total})",
                anchor='w',
                fill=self.fg_color,
                font=("Arial", 10)  # Bold for emphasis
            )

        except Exception as e:
            print(f"Error drawing pie chart: {e}")

    def filter_by_status(self, status):
        print(f"[DEBUG] Filtering applications by status: {status}")
        filtered_df = self.applications_df[self.applications_df['Status'] == status]
        self.populate_treeview(filtered_df)

    def draw_submissions_line_graph(self, parent_frame):
        """
        Draws a line graph representing the number of applications submitted over time.
        The x-axis represents the dates (formatted to show only months), and the y-axis
        represents the number of applications.

        Parameters:
        - parent_frame (tk.Frame): The frame where the line graph Canvas will be placed.
        """
        try:
            # Validate color attributes with defaults
            bg_color = getattr(self, 'bg_color', '#FFFFFF')
            entry_bg_color = getattr(self, 'entry_bg_color', '#F0F0F0')
            fg_color = getattr(self, 'fg_color', '#000000')

            # Check if 'Date Applied' column exists
            if 'Date Applied' not in self.applications_df.columns:
                return

            # Ensure 'Date Applied' is in datetime format
            self.applications_df['Date Applied'] = pd.to_datetime(self.applications_df['Date Applied'], errors='coerce')

            # Drop rows with invalid dates
            valid_dates_df = self.applications_df.dropna(subset=['Date Applied'])

            if valid_dates_df.empty:
                return

            # Group by date and count the number of applications per date
            submissions_per_date = (
                valid_dates_df.groupby(valid_dates_df['Date Applied'].dt.to_period('D'))
                .size()
                .reset_index(name='Count')
            )
            submissions_per_date['Date Applied'] = submissions_per_date['Date Applied'].dt.to_timestamp()

            # Sort the data by date
            submissions_per_date = submissions_per_date.sort_values('Date Applied')

            # Create a Matplotlib figure with reduced size
            fig, ax = plt.subplots(figsize=(4, 2), facecolor=bg_color)  # Increased size for better visibility

            # Plot the line graph with reduced marker size
            ax.plot(submissions_per_date['Date Applied'], submissions_per_date['Count'],
                    marker='o', markersize=3, linestyle='-', color='#9fc5e8')  # Using a theme-consistent color

            # Set the facecolor of the axes
            ax.set_facecolor(entry_bg_color)

            # Set tick params with theme colors and reduced label sizes
            ax.tick_params(axis='x', colors=fg_color, labelsize=8)
            ax.tick_params(axis='y', colors=fg_color, labelsize=8)

            # Format the x-axis dates to show only the month names
            ax.xaxis.set_major_locator(mdates.MonthLocator())  # Set major ticks to every month
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))  # Format ticks as abbreviated month names

            # Rotate x-axis labels if necessary
            fig.autofmt_xdate(rotation=0, ha='right')

            # Remove the top and right spines for a cleaner look
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_visible(False)
            ax.spines['bottom'].set_visible(False)

            # Adjust layout for tightness
            plt.tight_layout()

            # Destroy previous canvas and figure if they exist
            if hasattr(self, 'line_graph_canvas') and self.line_graph_canvas:
                self.line_graph_canvas.get_tk_widget().destroy()
                plt.close(self.line_graph_canvas.figure)

            # Embed the Matplotlib figure into Tkinter
            canvas = FigureCanvasTkAgg(fig, master=parent_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)

            # Store the canvas to prevent garbage collection
            self.line_graph_canvas = canvas


        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while drawing the line graph:\n{e}")

    # Theme Change
    def set_theme(self, theme_name):
        """
        Set the application's theme to either 'Light' or 'Dark'.

        Parameters:
        - theme_name (str): The name of the theme to apply ('Light' or 'Dark').
        """
        print(f"[DEBUG] Setting theme to {theme_name} mode.")
        if theme_name not in ["Light", "Dark"]:
            messagebox.showerror("Invalid Theme", "Selected theme is not supported.")
            return

        # Apply the selected theme
        if theme_name == "Light":
            self.set_light_mode()
        elif theme_name == "Dark":
            self.set_dark_mode()

        # Update the theme attribute
        self.theme = theme_name

        # Persist the theme choice to the configuration
        self.update_config(theme=theme_name)
        print(f"[DEBUG] Theme set to {theme_name} and saved to configuration.")

        # Refresh the GUI to apply changes
        self.apply_theme()

    # Treeview Setup and Interaction
    def _on_mousewheel_clipboard(self, event):
        """Handle mouse wheel scrolling for the Clipboard tab."""
        if sys.platform.startswith('win'):
            # For Windows, event.delta is a multiple of 120
            self.clipboard_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif sys.platform.startswith('darwin'):
            # For macOS, event.delta is also a multiple of 120 but may require different handling
            self.clipboard_canvas.yview_scroll(int(-1 * (event.delta)), "units")
        else:
            # For Linux, event.num is 4 (scroll up) or 5 (scroll down)
            if event.num == 4:
                self.clipboard_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.clipboard_canvas.yview_scroll(1, "units")

    def populate_treeview(self, df):
        """
        Populates the Treeview with data from the provided DataFrame.
        Each Treeview item ID corresponds to the DataFrame's index.
        """
        # Clear existing data
        for item in self.applications_tree.get_children():
            self.applications_tree.delete(item)

        # Insert new data
        for index, row in df.iterrows():
            # Use the DataFrame index as the Treeview item ID
            self.applications_tree.insert(
                "",
                "end",
                iid=str(index),  # Ensure item IDs are strings
                values=(
                    index + 1,  # Assuming 'No' starts at 1
                    row.get("Company", ""),
                    row.get("Position", ""),
                    row.get("Application Portal URL", ""),
                    row.get("Date Applied", ""),
                    row.get("Status", "")
                )
            )

            print(f"[DEBUG] Populating Treeview with {len(df)} records.")
            pass  # Replace with your actual Treeview population logic

    def refresh_treeview(self):
        """
        Clears and repopulates the Treeview with the most current DataFrame Data.
        """
        # Clear existing Treeview Data
        for item in self.applications_tree.get_children():
            self.applications_tree.delete(item)

        # Reload and display the current Data in the Treeview
        self.populate_treeview(self.applications_df)

    def on_treeview_click(self, event):
        """
        Handles single-click events on the Treeview cells.
        Supports dropdown selection for 'Status' column only.
        """
        # Close any open status dropdown
        if self.status_combobox:
            self.status_combobox.destroy()
            self.status_combobox = None

        # Identify the selected row and column
        selected_item = self.applications_tree.selection()
        if not selected_item:
            return  # Exit if no item is selected

        selected_item = selected_item[0]
        column = self.applications_tree.identify_column(event.x)
        col_index = int(column.replace("#", "")) - 1  # Convert column to zero-based index

        # Show status dropdown if the "Status" column is clicked
        if col_index == 5:
            self.show_status_dropdown(selected_item, col_index)
            return

        # No action for other columns on single click
        # Remove selection tracking to prevent interference
        # self.selected_row = None
        # self.selected_column = None

    def on_treeview_double_click(self, event):
        """
        Handles double-click events on the Treeview.
        If the double-click is on the 'Application Portal URL' column, opens the URL in the default web browser.
        """
        # Identify the region where the double-click occurred
        region = self.applications_tree.identify("region", event.x, event.y)
        if region != "cell":
            print("[DEBUG] Double-click occurred outside a cell region.")
            return  # Only proceed if double-clicked on a cell

        # Get the row and column indices
        row_id = self.applications_tree.identify_row(event.y)
        column_id = self.applications_tree.identify_column(event.x)
        col_index = int(column_id.replace("#", "")) - 1  # Convert to zero-based index

        # Log the identified row and column for debugging
        print(f"[DEBUG] Double-click detected on row_id: {row_id}, column_id: {column_id} (col_index: {col_index})")

        # Define the index of the 'Application Portal URL' column
        try:
            url_col_index = self.applications_tree["columns"].index("Application Portal URL")
        except ValueError:
            print("[ERROR] 'Application Portal URL' column not found.")
            logging.error("Error: 'Application Portal URL' column not found.")
            return

        # Check if the double-click was on the 'Application Portal URL' column
        if col_index == url_col_index:
            # Retrieve the URL from the Treeview
            item_values = self.applications_tree.item(row_id, "values")
            if not item_values:
                print("[DEBUG] No values found for the selected row.")
                return

            url = item_values[col_index]
            print(f"[DEBUG] URL identified for opening: {url}")

            if url and (url.startswith("http://") or url.startswith("https://")):
                try:
                    webbrowser.open(url, new=2)  # Open in a new browser tab
                    print(f"[INFO] Opening URL: {url}")
                except Exception as e:
                    print(f"[ERROR] Could not open URL: {url}. Exception: {e}")
                    logging.error(f"Error opening URL '{url}': {e}")
            else:
                print("[DEBUG] Invalid or missing URL.")
                messagebox.showerror("Invalid URL", "The selected application does not have a valid URL.")
        else:
            print("[DEBUG] Double-click detected on a non-URL column.")

    def on_treeview_cell_edit(self, event):
        """
        Enables direct in-cell editing for non-URL, non-Status columns.
        Activates an editable entry when the relevant cell is selected.
        """
        selected_item = self.applications_tree.selection()
        if not selected_item:
            return  # Exit if no item is selected

        selected_item = selected_item[0]
        column = self.applications_tree.identify_column(event.x)
        col_index = int(column[1:]) - 1  # Zero-based index in applications_df

        # Allow editing for non-Status and non-URL columns only
        if self.applications_tree["columns"][col_index] not in ["Status", "Application Portal URL"]:
            self.applications_tree.bind("<KeyRelease>", lambda e: self.save_direct_edit(selected_item, col_index))

    def sort_treeview_column(self, col, reverse):
        """
        Sorts the Treeview column data in ascending or descending order.

        Parameters:
        - col (str): The column name to sort.
        - reverse (bool): Whether to reverse the sorting order.
        """
        try:
            # Extract data from the Treeview for sorting
            data = [
                (self._convert_to_sortable(self.applications_tree.set(child, col)), child)
                for child in self.applications_tree.get_children("")
            ]

            # Sort data
            data.sort(key=lambda x: x[0], reverse=reverse)

            # Reorder the Treeview items based on sorted data
            for index, (_, child) in enumerate(data):
                self.applications_tree.move(child, "", index)

            # Update column header with reversed sorting order for subsequent clicks
            self.applications_tree.heading(
                col, command=lambda: self.sort_treeview_column(col, not reverse)
            )
        except Exception as e:
            print(f"[ERROR] Sorting failed for column '{col}': {e}")

    def _convert_to_sortable(self, value):
        """
        Converts a value to a sortable type, prioritizing numeric sorting.

        Parameters:
        - value (str): The value to convert.

        Returns:
        - A sortable type (int, float, or string).
        """
        if value == "" or value is None:
            return ""  # Treat empty values as empty strings
        try:
            # Attempt to convert to an integer
            return int(value)
        except ValueError:
            try:
                # Attempt to convert to a float if not an integer
                return float(value)
            except ValueError:
                # Default to string if not numeric
                return value

    # Editing and Saving
    def mask_value(self, value, mask_char="*"):
        """
        Returns a masked version of the input value using the specified mask character.

        Parameters:
        - value (str): The string to be masked.
        - mask_char (str): The character to use for masking.

        Returns:
        - str: The masked string.
        """
        return mask_char * len(value) if value else ""

    def toggle_mask(self, label):
        """
        Toggles the masking of a specific personal information entry.

        Parameters:
        - label (str): The label of the personal information entry to toggle.
        """
        # Ensure self.mask_vars exists
        if not hasattr(self, 'mask_vars'):
            self.mask_vars = {}

        mask_var = self.mask_vars.get(label)
        if not mask_var:
            return

        # Load current personal info
        personal_info = load_personal_info()
        actual_value = personal_info.get(label, {}).get("value", "")
        current_masked = mask_var.get()

        if current_masked:
            # Mask the value
            masked_value = self.mask_value(actual_value)
            # Update the main window's label
            info_label = self.personal_info_entries.get(label)
            if info_label:
                info_label.config(text=masked_value, foreground="#AAAAAA")
        else:
            # Show the actual value
            info_label = self.personal_info_entries.get(label)
            if info_label:
                info_label.config(text=actual_value, foreground=self.fg_color)

        # Save the masking flag
        self.save_personal_info()

    def get_actual_personal_info_value(self, label):
        """
        Retrieves the actual value of a personal information entry from the JSON file.

        Parameters:
        - label (str): The label of the personal information entry.

        Returns:
        - str: The actual value of the entry.
        """
        personal_info = load_personal_info()
        entry = personal_info.get(label, {})
        return entry.get("value", "")

    def create_edit_entry(self, item_id, col_index):
        """
        Creates an Entry widget directly within the Treeview cell, allowing in-cell editing.
        Only available for non-URL columns.
        """
        # Get the bounding box coordinates of the cell for placing the Entry widget
        x, y, width, height = self.applications_tree.bbox(item_id, column="#" + str(col_index + 1))

        # Retrieve the current value of the cell for editing
        current_value = self.applications_tree.item(item_id, "values")[col_index]

        # Create and place an Entry widget at the cell's location with its current value
        self.edit_entry = tk.Entry(self.applications_tree, width=width)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.place(x=x, y=y, width=width, height=height)

        # Bind events to save or cancel edit on Enter key or focus loss
        self.edit_entry.bind("<Return>", lambda e: self.save_edit(item_id, col_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.edit_entry.destroy())
        self.edit_entry.focus_set()  # Set focus to the entry widget for immediate editing

    def save_edit(self, item_id, col_index):
        """
        Saves the edited value from the Entry widget back to both the Treeview cell and the DataFrame.
        """
        # Check if edit_entry exists and retrieve the new value
        if self.edit_entry:
            new_value = self.edit_entry.get()

            # Update the Treeview cell with the new value
            values = list(self.applications_tree.item(item_id, "values"))
            values[col_index] = new_value
            self.applications_tree.item(item_id, values=values)

            # Update the DataFrame with the new value
            column_name = self.applications_tree["columns"][col_index]
            self.applications_df.at[int(item_id), column_name] = new_value

            # Save the DataFrame to Excel
            save_applications_to_excel(self.applications_df)

            # Conditionally sync updated data to Google Sheets if sync is enabled
            if self.sync_to_google:
                try:
                    write_to_google_sheets(self.applications_df)
                    print(f"Updated '{column_name}' synced to Google Sheets for row {item_id}.")
                except Exception as e:
                    print(f"Error: Could not sync with Google Sheets. {str(e)}")
            else:
                print("Google Sync is disabled. Changes were not synced to Google Sheets.")

            # Destroy the Entry widget after saving the edit
            self.edit_entry.destroy()
            self.edit_entry = None  # Reset edit_entry

            print(f"Saved edit: {new_value} in cell ({item_id}, {col_index}).")
        else:
            print("Edit entry does not exist to save.")

    def save_direct_edit(self, item_id, col_index):
        """
        Saves the edited value directly within the cell without needing an Entry widget.
        Only allows saving for non-URL and non-Status columns.
        """
        # Retrieve the directly edited value from the Treeview cell
        edited_value = self.applications_tree.item(item_id, "values")[col_index]

        # Update the DataFrame with the edited value for the specified column
        column_name = self.applications_tree["columns"][col_index]
        self.applications_df.at[int(item_id), column_name] = edited_value

        # Persist changes by saving the updated DataFrame to Excel
        save_applications_to_excel(self.applications_df)

        # Unbind the key release event after saving to prevent unintended edits
        self.applications_tree.unbind("<KeyRelease>")

    # Clipboard and Copying
    def copy_to_clipboard(self, value):
        """
        Copies the specified value to the system clipboard.

        Parameters:
        value (str): The text value to copy to the clipboard.
        """
        try:
            # Clear existing clipboard content
            self.clipboard_clear()

            # Append the specified value to the clipboard
            self.clipboard_append(value)

            # Force Tkinter to process the clipboard event immediately
            self.update()

            # Log confirmation of the copied value
            print(f"Copied to clipboard: {value}")
        except Exception as e:
            print(f"Error copying to clipboard: {e}")
            logging.error(f"Error copying to clipboard: {e}")

    def copy_rows(self, row_ids):
        """
        Copies the values of the selected rows from the Treeview to the clipboard.
        Each row's values are tab-separated, and rows are separated by newlines.
        """
        if not row_ids:
            print("No rows selected for copying.")
            return

        copied_text = ""
        for row_id in row_ids:
            # Retrieve all cell values for the specified row
            row_data = self.applications_tree.item(row_id, "values")
            # Concatenate row values into a single tab-separated string
            row_text = "\t".join(str(item) for item in row_data)
            copied_text += row_text + "\n"

        # Copy the concatenated text to the clipboard
        self.clipboard_clear()
        self.clipboard_append(copied_text.strip())

        # Log confirmation of the copied rows
        print(f"Copied {len(row_ids)} rows to clipboard.")

    # Data Management and Synchronization
    def sync_from_google_sheets(self):
        """Fetch data from Google Sheets if Google Sync is enabled."""
        if not self.sync_to_google:
            print("Google Sync is disabled. Skipping sync from Google Sheets.")
            return

        try:
            # Retrieve the latest data from Google Sheets
            google_df = read_from_google_sheets()

            # Replace NaN values with empty strings
            google_df = google_df.fillna('')

            # Check for differences and update if necessary
            if not google_df.empty:
                if not google_df.equals(self.applications_df):
                    print("Detected changes in Google Sheets. Updating local data.")
                    self.applications_df = google_df

                    # Ensure the Treeview is initialized before updating it
                    if hasattr(self, 'applications_tree') and self.applications_tree:
                        self.populate_treeview(self.applications_df)
                    else:
                        print("Error: applications_tree is not initialized yet. Will populate later.")

        except Exception as e:
            print(f"Error syncing data from Google Sheets: {e}")
            # Log the error but do not disable Google Sync
            logging.error(f"Error syncing data from Google Sheets: {e}")

    def sync_to_google_sheets(self):
        """Push local DataFrame data to Google Sheets if Google Sync is enabled."""
        if not self.sync_to_google:
            print("Google Sync is disabled. Skipping sync to Google Sheets.")
            return

        try:
            # Update Google Sheets with the current DataFrame data
            write_to_google_sheets(self.applications_df)
            print("Data synced to Google Sheets successfully.")
        except Exception as e:
            print(f"Error syncing data to Google Sheets: {e}")
            # Log the error but do not disable Google Sync
            logging.error(f"Error syncing data to Google Sheets: {e}")

    def schedule_sync(self):
        """Schedules periodic syncing from Google Sheets every 60 seconds."""
        # Schedule the next sync and store the task ID
        self.sync_task = self.after(60000, self.schedule_sync)

        if self.sync_to_google:
            # Perform synchronization with Google Sheets
            self.sync_from_google_sheets()

    def save_application(self):
        """
        Captures Data from input fields, validates it, and saves it as a new application entry.
        Updates both the local DataFrame and Google Sheets, then refreshes the Treeview and graphs.
        """
        # Retrieve input Data and clean up extra spaces
        position = self.position_entry.get().strip()
        company = self.company_entry.get().strip()
        url = self.url_entry.get().strip()  # URL is optional
        date_applied = datetime.now().strftime("%Y-%m-%d")
        status = "Submitted"  # Default status for new applications

        # Validate required fields (company and position)
        if not company or not position:
            print("Please fill out the Company and Position fields before adding an application.")
            messagebox.showerror("Error", "Company and Position are required fields.")
            return  # Stop if required fields are missing

        # Ensure DataFrame has the correct columns if it's empty
        if self.applications_df.empty:
            self.applications_df = pd.DataFrame(
                columns=["Company", "Position", "Application Portal URL", "Date Applied", "Status"])

        # Log current DataFrame columns for debugging purposes
        print("Current DataFrame columns:", self.applications_df.columns)

        # Create a new row of Data in DataFrame format
        new_data = pd.DataFrame(
            [[company, position, url, date_applied, status]],
            columns=["Company", "Position", "Application Portal URL", "Date Applied", "Status"]
        )

        # Append the new Data to the applications DataFrame
        self.applications_df = pd.concat([self.applications_df, new_data], ignore_index=True)

        # Save the updated DataFrame to the local Excel file
        save_applications_to_excel(self.applications_df)
        print("Data saved locally to Excel.")

        # Sync updated Data to Google Sheets only if sync is enabled
        if self.sync_to_google:
            try:
                write_to_google_sheets(self.applications_df)
                print("Data synced to Google Sheets.")
            except FileNotFoundError as e:
                print(f"Google Sheets sync failed: {e}")
                messagebox.showerror("Error", f"Google Sheets sync failed: {e}")
            except Exception as e:
                print(f"[ERROR] Unexpected error during Google Sheets sync: {e}")
                messagebox.showerror("Error", f"Google Sheets sync failed: {e}")

        # Refresh the Treeview to display the new application
        self.populate_treeview(self.applications_df)

        # Redraw the pie chart within the pie_chart_frame
        if hasattr(self, 'pie_chart_frame'):
            self.draw_status_pie_chart(self.pie_chart_frame)
        else:
            print("[DEBUG] pie_chart_frame not found. Cannot redraw pie chart.")

        # Redraw the line graph within the line_graph_frame
        if hasattr(self, 'line_graph_frame'):
            self.draw_submissions_line_graph(self.line_graph_frame)
        else:
            print("[DEBUG] line_graph_frame not found. Cannot redraw line graph.")

        # Clear the input fields after saving
        self.clear_input_fields()

    def clear_input_fields(self):
        """
        Clears the input fields in the 'Add Application' tab.
        """
        self.company_entry.delete(0, tk.END)
        self.position_entry.delete(0, tk.END)
        self.url_entry.delete(0, tk.END)

    # Search and Filter
    def perform_search(self):
        """
        Filters the Treeview to display only rows containing the search term.
        If no search term is entered, all rows are displayed.
        """
        # Retrieve and clean the search term (convert to lowercase for case-insensitive matching)
        search_term = self.search_var.get().strip().lower()

        # If the search term is empty, display all rows
        if not search_term:
            self.populate_treeview(self.applications_df)
            return

        # Filter the DataFrame: retain rows that contain the search term in any column
        filtered_df = self.applications_df[
            self.applications_df.apply(
                lambda row: search_term in row.astype(str).str.lower().to_string(), axis=1
            )
        ]

        # Refresh the Treeview to show only the rows in the filtered DataFrame
        self.populate_treeview(filtered_df)

    # Context Menu and Cell Interaction
    def show_context_menu(self, event):
        """
        Displays a context menu at the cursor's current position.
        """
        # Identify the row and column where the click occurred
        row_id = self.applications_tree.identify_row(event.y)
        column_id = self.applications_tree.identify_column(event.x)
        col_index = int(column_id.replace("#", "")) - 1  # Convert to zero-based index

        # Get all selected rows
        selected_rows = self.applications_tree.selection()

        # Only show the context menu if a row is clicked
        if not row_id:
            return

        # Create the context menu
        context_menu = tk.Menu(self, tearoff=0)

        if len(selected_rows) > 1:
            # If multiple rows are selected, provide the option to delete all
            context_menu.add_command(label="Delete Selected Rows", command=lambda: self.delete_rows(selected_rows))
            context_menu.add_command(label="Copy Selected Rows", command=lambda: self.copy_rows(selected_rows))
        else:
            # General options: Delete Row, Copy Row
            context_menu.add_command(label="Delete Row", command=lambda: self.delete_rows([row_id]))
            context_menu.add_command(label="Copy Row", command=lambda: self.copy_row(row_id))

            # Column-specific options based on the column index
            if col_index == 1:  # Company column
                context_menu.add_command(label="Edit Company",
                                         command=lambda: self.edit_cell(row_id, col_index, "Company"))
            elif col_index == 2:  # Position column
                context_menu.add_command(label="Edit Position",
                                         command=lambda: self.edit_cell(row_id, col_index, "Position"))
            elif col_index == 3:  # URL Portal column
                context_menu.add_command(label="Edit URL",
                                         command=lambda: self.edit_cell(row_id, col_index, "Application Portal URL"))
            elif col_index == 4:  # Date Applied column
                context_menu.add_command(label="Edit Date",
                                         command=lambda: self.edit_cell(row_id, col_index, "Date Applied"))
            elif col_index == 5:  # Status column
                context_menu.add_command(label="Edit Status",
                                         command=lambda: self.show_status_dropdown(row_id, col_index))

        # Display the context menu at the mouse cursor position
        context_menu.tk_popup(event.x_root, event.y_root)

    def delete_rows(self, row_ids):
        """
        Deletes the selected rows from the Treeview, DataFrame, and Google Sheets.
        Updates the local Excel file and Treeview to reflect the deletion.
        """
        if not row_ids:
            print("No rows selected for deletion.")
            return

        # Convert row_ids to integers and sort them in descending order to prevent index shifting
        row_indices = sorted([int(row_id) for row_id in row_ids], reverse=True)

        # Remove the rows from the DataFrame and Treeview
        for row_index in row_indices:
            if row_index in self.applications_df.index:
                # Remove the row from the DataFrame
                self.applications_df = self.applications_df.drop(row_index)
            else:
                print(f"Row ID {row_index} not found in DataFrame index.")

        # Reset the DataFrame index after deletions
        self.applications_df.reset_index(drop=True, inplace=True)

        # Update the Treeview
        self.populate_treeview(self.applications_df)

        # Sync with Google Sheets if enabled
        if self.sync_to_google:
            try:
                # Update Google Sheets with the current DataFrame data
                write_to_google_sheets(self.applications_df)
                print("Data synced to Google Sheets after deletion.")
            except Exception as e:
                print(f"Error syncing data to Google Sheets: {e}")
        else:
            print("Google Sync is disabled. Changes were not synced to Google Sheets.")

        # Save the updated DataFrame to Excel
        save_applications_to_excel(self.applications_df)
        print("Data saved locally to Excel after deletion.")
        # After updating status in save_edit or save_application
        self.draw_status_pie_chart(self.pie_chart_frame)  # Ensure correct frame reference

    def edit_cell(self, row_id, col_index, column_name):
        """
        Creates an Entry widget directly over the specified Treeview cell for inline editing.
        """
        # If an edit Entry already exists, destroy it to prevent multiple editors
        if self.edit_entry:
            self.edit_entry.destroy()

        # Retrieve the cell coordinates and current value
        x, y, width, height = self.applications_tree.bbox(row_id, column="#" + str(col_index + 1))
        current_value = self.applications_tree.item(row_id, "values")[col_index]

        # Create an Entry widget for editing
        self.edit_entry = tk.Entry(self.applications_tree, width=width)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus_set()  # Ensure the Entry widget is focused

        # Bind actions to save on Enter key press or focus out
        self.edit_entry.bind("<Return>", lambda e: self.save_edit(row_id, col_index))
        self.edit_entry.bind("<FocusOut>", lambda e: self.save_edit(row_id, col_index))

    def show_status_dropdown(self, item_id, col_index):
        """
        Displays a dropdown menu for editing the 'Status' column in the Treeview.

        Parameters:
        - item_id (str): Identifier of the row containing the status to edit.
        - col_index (int): Index of the 'Status' column.
        """
        status_options = ["Submitted", "Rejected", "Interview", "Offer"]
        current_status = self.applications_tree.item(item_id, "values")[col_index]

        # Create a dropdown menu (Combobox) with status options
        self.status_combobox = ttk.Combobox(self.applications_tree, values=status_options, state="readonly")
        self.status_combobox.set(current_status)
        x, y, width, height = self.applications_tree.bbox(item_id, column="#" + str(col_index + 1))

        # Position the dropdown if coordinates are valid
        if x and y:
            self.status_combobox.place(x=x, y=y, width=width, height=height)
        else:
            print("Error: Unable to place the combobox due to invalid bounding box values.")

        # Focus on the dropdown and bind selection event for saving
        self.status_combobox.focus_set()
        self.status_combobox.bind("<<ComboboxSelected>>", lambda event: self.save_status(item_id, col_index))

    def save_status(self, item_id, col_index):
        """
        Saves the selected status from the dropdown to the Treeview, DataFrame, and Google Sheets.
        """
        # Retrieve the new status from the dropdown menu
        new_status = self.status_combobox.get()
        values = list(self.applications_tree.item(item_id, "values"))
        values[col_index] = new_status
        self.applications_tree.item(item_id, values=values)

        # Update the DataFrame with the new status
        column_name = self.applications_tree["columns"][col_index]
        self.applications_df.at[int(item_id), column_name] = new_status

        # Save changes to the Excel file locally
        try:
            save_applications_to_excel(self.applications_df, DATA_FILE_PATH)
            print(f"Status '{new_status}' saved for row {item_id} in Excel.")
        except Exception as e:
            print(f"Error: Could not save to the Excel file. {str(e)}")

        # Conditionally sync the updated status to Google Sheets if sync is enabled
        if self.sync_to_google:
            try:
                write_to_google_sheets(self.applications_df)
                print(f"Status '{new_status}' synced with Google Sheets for row {item_id}.")
            except Exception as e:
                print(f"Error: Could not sync with Google Sheets. {str(e)}")
        else:
            print("Google Sync is disabled. Changes were not synced to Google Sheets.")

        # After updating status in save_edit or save_application
        self.draw_status_pie_chart(self.pie_chart_frame)  # Ensure correct frame reference

        # Destroy the dropdown after saving
        self.status_combobox.destroy()
        self.status_combobox = None

    # Personal Info
    def save_personal_info(self):
        """Save the masking flags to personal_info.json."""
        personal_info = load_personal_info()
        for label, mask_var in self.mask_vars.items():
            if label in personal_info:
                personal_info[label]['masked'] = mask_var.get()
        try:
            with open(PERSONAL_INFO_FILE, "w") as file:
                json.dump(personal_info, file, indent=4)
            print("Personal information masking flags updated successfully.")
        except Exception as e:
            print(f"Error saving personal information masking flags: {e}")
            messagebox.showerror("Error", f"Failed to save personal information masking flags: {e}")

    # Setup Settings
    def start_move(self, event):
        self.xwin = self.winfo_x()
        self.ywin = self.winfo_y()
        self.startx = event.x_root
        self.starty = event.y_root

    def do_move(self, event):
        deltax = event.x_root - self.startx
        deltay = event.y_root - self.starty
        x = self.xwin + deltax
        y = self.ywin + deltay
        self.geometry(f"+{x}+{y}")

    def toggle_settings_menu(self, event=None):
        """Toggle the visibility of the settings menu."""
        if self.menu_visible:
            self.settings_menu.unpost()  # Hide the menu if it is already open
        else:
            self.settings_menu.post(self.settings_button.winfo_rootx(),
                                    self.settings_button.winfo_rooty() + self.settings_button.winfo_height())
        self.menu_visible = not self.menu_visible  # Toggle the visibility state

    def toggle_sync(self):
        """Toggle the Google Sync feature based on the Checkbutton state."""
        self.sync_to_google = self.google_sync_var.get()
        print(f"Sync to Google Sheets: {'Enabled' if self.sync_to_google else 'Disabled'}")

        # Update configuration: Only ENABLE_GOOGLE_SYNC
        self.update_config(ENABLE_GOOGLE_SYNC=self.sync_to_google)

        # Cancel any existing scheduled sync if disabling
        if not self.sync_to_google and self.sync_task is not None:
            self.after_cancel(self.sync_task)
            self.sync_task = None
            print("Scheduled Google Sync tasks canceled.")

        # Schedule sync tasks if enabling
        if self.sync_to_google:
            self.schedule_sync()
            print("Scheduled Google Sync tasks.")

        # Optionally, perform an immediate sync when enabled
        if self.sync_to_google:
            self.sync_to_google_sheets()

    def apply_theme(self):
        """Apply the selected theme to all widgets."""
        print(f"[DEBUG] Applying {self.theme} theme to all widgets.")
        # Update all widgets
        self.update_all_widgets_theme(self)

        # Update menu bar
        self.update_menu_bar_theme()

        # Update Entry widget cursor colors and background/foreground
        self.update_entry_widgets()

        # Redraw the pie chart to update legend labels with new theme colors
        if hasattr(self, 'pie_chart_frame'):
            self.draw_status_pie_chart(self.pie_chart_frame)

        # Redraw the line graph to update its appearance based on the new theme
        if hasattr(self, 'line_graph_frame'):
            self.draw_submissions_line_graph(self.line_graph_frame)

        # Refresh the GUI to apply changes
        self.update_idletasks()

        print(f"[DEBUG] {self.theme} theme applied to all widgets.")

    def update_entry_widgets(self):
        """
        Update the background and foreground colors for all Entry widgets based on the current theme.
        """
        for entry in self.entry_widgets:
            try:
                entry.config(bg=self.entry_bg_color, fg=self.entry_fg_color, insertbackground=self.fg_color)
            except tk.TclError as e:
                print(f"[ERROR] Failed to update Entry widget colors: {e}")
                logging.error(f"Failed to update Entry widget colors: {e}")

    def set_light_mode(self):
        """Apply Light Mode theme settings."""
        print("[DEBUG] Applying Light Mode theme.")
        self.bg_color = "#F0F0F0"
        self.fg_color = "#000000"
        self.entry_bg_color = "#FFFFFF"
        self.entry_fg_color = "#000000"
        self.button_bg_color = "#E0E0E0"
        self.menu_bg_color = "#E0E0E0"
        self.menu_fg_color = "#000000"
        self.menu_active_bg = "#C0C0C0"

        style = ttk.Style()
        style.theme_use("alt")  # Use 'alt' theme for better customization

        # Configure styles for ttk widgets
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color)
        style.configure("TFrame", background=self.bg_color)
        style.configure("TButton", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TEntry", fieldbackground=self.entry_bg_color, foreground=self.entry_fg_color)
        style.configure("Treeview", background=self.entry_bg_color, foreground=self.entry_fg_color,
                        fieldbackground=self.entry_bg_color)
        style.map('Treeview', background=[('selected', '#D9D9D9')], foreground=[('selected', '#000000')])
        style.configure("Treeview.Heading", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TNotebook", background=self.bg_color)
        style.configure("TNotebook.Tab", background=self.bg_color, foreground=self.fg_color)
        style.map("TNotebook.Tab", background=[('selected', self.entry_bg_color)])
        style.configure("TCombobox", fieldbackground=self.entry_bg_color, background=self.entry_bg_color,
                        foreground=self.entry_fg_color)
        style.map('TCombobox', fieldbackground=[('readonly', self.entry_bg_color)],
                  background=[('readonly', self.entry_bg_color)],
                  foreground=[('readonly', self.entry_fg_color)])

        # Custom style for Save buttons
        style.configure(
            "Custom.TButton",
            background=self.button_bg_color,
            foreground=self.fg_color,
            borderwidth=1,
            focusthickness=3,
            focuscolor='none',
            font=('TkDefaultFont', 12),
            padding=(10, 5)
        )
        style.map(
            "Custom.TButton",
            background=[('active', self.menu_active_bg)],
            foreground=[('active', self.fg_color)]
        )

        # Update the Checkbutton's colors to match light mode
        self.google_sync_checkbutton.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            selectcolor=self.menu_bg_color
        )

        print("[DEBUG] Light Mode theme applied.")

    def set_dark_mode(self):
        """Apply Dark Mode theme settings."""
        print("[DEBUG] Applying Dark Mode theme.")
        self.bg_color = "#2E2E2E"
        self.fg_color = "#FFFFFF"
        self.entry_bg_color = "#3A3A3A"
        self.entry_fg_color = "#FFFFFF"
        self.button_bg_color = "#3E3E3E"
        self.menu_bg_color = "#3E3E3E"
        self.menu_fg_color = "#FFFFFF"
        self.menu_active_bg = "#5E5E5E"

        style = ttk.Style()
        style.theme_use("alt")  # Use 'alt' theme for better customization

        # Configure styles for ttk widgets
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color)
        style.configure("TFrame", background=self.bg_color)
        style.configure("TButton", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TEntry", fieldbackground=self.entry_bg_color, foreground=self.entry_fg_color)
        style.configure("Treeview", background=self.entry_bg_color, foreground=self.entry_fg_color,
                        fieldbackground=self.entry_bg_color)
        style.map('Treeview', background=[('selected', '#6A6A6A')], foreground=[('selected', '#FFFFFF')])
        style.configure("Treeview.Heading", background=self.button_bg_color, foreground=self.fg_color)
        style.configure("TNotebook", background=self.bg_color)
        style.configure("TNotebook.Tab", background=self.bg_color, foreground=self.fg_color)
        style.map("TNotebook.Tab", background=[('selected', self.entry_bg_color)])
        style.configure("TCombobox", fieldbackground=self.entry_bg_color, background=self.entry_bg_color,
                        foreground=self.entry_fg_color)
        style.map('TCombobox', fieldbackground=[('readonly', self.entry_bg_color)],
                  background=[('readonly', self.entry_bg_color)],
                  foreground=[('readonly', self.entry_fg_color)])

        # Custom style for Save buttons
        style.configure(
            "Custom.TButton",
            background=self.button_bg_color,
            foreground=self.fg_color,
            borderwidth=1,
            focusthickness=3,
            focuscolor='none',
            font=('TkDefaultFont', 12),
            padding=(10, 5)
        )
        style.map(
            "Custom.TButton",
            background=[('active', self.menu_active_bg)],
            foreground=[('active', self.fg_color)]
        )

        # Update the Checkbutton's colors to match dark mode
        self.google_sync_checkbutton.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            selectcolor=self.menu_bg_color
        )

        print("[DEBUG] Dark Mode theme applied.")

    def update_all_widgets_theme(self, widget):
        """Recursively update the theme for all widgets."""
        for child in widget.winfo_children():
            # Check if the widget is a ttk widget
            if isinstance(child, ttk.Widget):
                pass  # Skip ttk widgets; they're styled via ttk.Style
            else:
                # Get the list of options the widget supports
                options = child.keys()
                # Set 'bg' or 'background' if supported
                if 'bg' in options or 'background' in options:
                    child.config(bg=self.bg_color)
                # Set 'fg' or 'foreground' if supported
                if 'fg' in options or 'foreground' in options:
                    child.config(fg=self.fg_color)

                # Specific adjustments for certain widget types
                if isinstance(child, tk.Entry):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.entry_bg_color)
                    if 'fg' in options or 'foreground' in options:
                        child.config(fg=self.entry_fg_color)
                elif isinstance(child, tk.Button):
                    if 'activebackground' in options:
                        child.config(activebackground=self.button_bg_color)
                    if 'activeforeground' in options:
                        child.config(activeforeground=self.fg_color)
                elif isinstance(child, tk.Text):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.entry_bg_color)
                    if 'fg' in options or 'foreground' in options:
                        child.config(fg=self.entry_fg_color)
                # For frames and toplevels, only set 'bg'
                elif isinstance(child, (tk.Frame, tk.Toplevel)):
                    if 'bg' in options or 'background' in options:
                        child.config(bg=self.bg_color)
            # Recursively update child widgets
            self.update_all_widgets_theme(child)

    def update_menu_bar_theme(self):
        """Update theme for the menu bar components (Settings and Google Sync Checkbutton)."""
        # Update menu bar background color
        self.menu_bar.config(bg=self.menu_bg_color)

        # Update Styles for Settings.TButton
        style = ttk.Style()
        style.configure("Settings.TButton",
                        background=self.menu_bg_color,
                        foreground=self.menu_fg_color,
                        font=("Arial", 12),
                        relief="flat")

        style.map("Settings.TButton",
                  background=[('active', self.menu_active_bg)],
                  foreground=[('active', self.menu_fg_color)])

        # Update the Checkbutton's colors based on theme
        self.google_sync_checkbutton.config(
            bg=self.menu_bg_color,
            fg=self.menu_fg_color,
            activebackground=self.menu_active_bg,
            activeforeground=self.menu_fg_color,
            selectcolor=self.menu_bg_color
        )

    def update_entry_cursor_colors(self):
        """
        Update the cursor color for all Entry widgets based on the current theme.
        """
        cursor_color = self.fg_color
        for entry in self.entry_widgets:
            try:
                entry.config(insertbackground=cursor_color)
            except tk.TclError as e:
                print(f"[ERROR] Failed to set insertbackground for an Entry widget: {e}")
                logging.error(f"Failed to set insertbackground for an Entry widget: {e}")

    def toggle_theme(self):
        """Toggle between Dark and Light themes."""
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()  # Correctly apply the selected theme
        # Save the theme to the configuration
        save_theme("Dark" if self.is_dark_mode else "Light")

        # Optionally, recreate or refresh all tabs to apply the new theme
        # If all widgets are updated via traversal, this might not be necessary
        # self.create_personal_info_tab()
        # self.create_add_application_tab()
        # self.create_view_edit_applications_tab()

    def _bind_mousewheel_events(self, widget, handler):
        """
        Binds mouse wheel events to a given widget for scrolling.

        Parameters:
        - widget: The widget to bind the events to (e.g., Canvas).
        - handler: The method to handle the scrolling.
        """
        if sys.platform.startswith('win'):
            widget.bind("<MouseWheel>", handler)
        elif sys.platform.startswith('darwin'):
            widget.bind("<MouseWheel>", handler)
        else:
            # Linux typically uses Button-4 (scroll up) and Button-5 (scroll down)
            widget.bind("<Button-4>", handler)
            widget.bind("<Button-5>", handler)

    def bind_events_to_children(self, parent_widget, click_handler, drop_handler=None):
        """
        Bind click and drag-and-drop events to all child widgets within the parent_widget,
        excluding interactive widgets.

        Parameters:
        - parent_widget: The frame whose child widgets will have events bound.
        - click_handler: The method to handle click events.
        - drop_handler: The method to handle drop events (optional).
        """
        widgets = parent_widget.winfo_children()
        for widget in widgets:
            # Exclude interactive widgets to allow normal user interaction
            if isinstance(widget, (tk.Entry, tk.Text, ttk.Entry, ttk.Combobox)):
                continue

            # Bind the click event
            widget.bind("<Button-1>", click_handler)

            # Register drag-and-drop if a drop_handler is provided
            if drop_handler:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', drop_handler)

            # Recursively bind events to child widgets
            self.bind_events_to_children(widget, click_handler, drop_handler)

    def select_app_file(self, event=None):
        """Prompt user to select Applications.xlsx, and copy it to AppData."""
        file_path = filedialog.askopenfilename(title="Select Applications.xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                shutil.copy(file_path, DATA_FILE_PATH)
                self.app_file_path_var.set(DATA_FILE_PATH)
                print(f"[DEBUG] Applications.xlsx copied to {DATA_FILE_PATH}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy Applications.xlsx: {e}")

    def select_service_account_file(self, event=None):
        """Prompt user to select Service Account JSON, and copy it to AppData."""
        file_path = filedialog.askopenfilename(title="Select Service Account JSON",
                                               filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                shutil.copy(file_path, SERVICE_ACCOUNT_FILE)
                self.service_account_file_path_var.set(SERVICE_ACCOUNT_FILE)
                print(f"[DEBUG] Service Account JSON copied to {SERVICE_ACCOUNT_FILE}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy Service Account JSON: {e}")

    def service_account_file_drop(self, event):
        """Handle the drop event for the Service Account JSON file."""
        print("[DEBUG] Service Account JSON file drop detected.")
        file_path = event.data
        file_list = self.tk.splitlist(file_path)
        if file_list:
            file_path = file_list[0]
            # Validate the file type
            if file_path.lower().endswith('.json'):
                try:
                    shutil.copy(file_path, SERVICE_ACCOUNT_FILE)
                    self.service_account_file_path_var.set(SERVICE_ACCOUNT_FILE)
                    print(f"[DEBUG] Service Account JSON copied to {SERVICE_ACCOUNT_FILE}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy Service Account JSON: {e}")
            else:
                messagebox.showerror("Invalid File", "Please drop a valid JSON file.")

    def app_file_drop(self, event):
        """Handle the drop event for the Applications.xlsx file."""
        print("[DEBUG] Applications.xlsx file drop detected.")
        file_path = event.data
        file_list = self.tk.splitlist(file_path)
        if file_list:
            file_path = file_list[0]
            # Validate the file type
            if file_path.lower().endswith('.xlsx'):
                try:
                    shutil.copy(file_path, DATA_FILE_PATH)
                    self.app_file_path_var.set(DATA_FILE_PATH)
                    print(f"[DEBUG] Applications.xlsx copied to {DATA_FILE_PATH}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy Applications.xlsx: {e}")
            else:
                messagebox.showerror("Invalid File", "Please drop a valid Excel (.xlsx) file.")

    def open_settings_dialog(self):
        """Open a dialog to configure the Service Account JSON and Spreadsheet ID."""
        dialog = tk.Toplevel(self)
        dialog.title("Google Sync Configuration")
        dialog.geometry("600x300")  # Adjusted size for larger content
        # Keep settings dialog in front
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus()

        # Use current theme colors
        bg = self.bg_color
        fg = self.fg_color
        entry_bg = self.entry_bg_color
        entry_fg = self.entry_fg_color
        button_bg = self.button_bg_color

        dialog.config(bg=bg)

        # --- Service Account JSON file ---
        # Fetch the latest Service Account File Path
        service_account_path = self.get_current_service_account_file_path()
        self.service_account_file_path_var = tk.StringVar(value=service_account_path)

        self.service_file_button = tk.Frame(dialog, bg=button_bg, relief='raised', bd=2)
        self.service_file_button.pack(pady=20, fill='x', padx=50)

        # Bind the entire frame for the service account button
        self.service_file_button.bind("<Enter>", lambda e: self.service_file_button.config(relief='groove'))
        self.service_file_button.bind("<Leave>", lambda e: self.service_file_button.config(relief='raised'))
        self.service_file_button.bind("<Button-1>", lambda e: self.select_service_account_file())

        # Use tkinterdnd2 methods for drag-and-drop
        self.service_file_button.drop_target_register(DND_FILES)
        self.service_file_button.dnd_bind('<<Drop>>', self.service_account_file_drop)

        # Place the icon and labels inside the frame for service account file
        self.service_icon_label = tk.Label(self.service_file_button, image=self.upload_json_icon, bg=button_bg)
        self.service_icon_label.pack(side='left', padx=(10, 5), pady=10)

        # Text label and path display for the service account file
        self.service_text_frame = tk.Frame(self.service_file_button, bg=button_bg)
        self.service_text_frame.pack(side='left', fill='x', expand=True)

        self.service_text_label = tk.Label(
            self.service_text_frame,
            text="Upload or Drop Service Account JSON File Here",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.service_text_label.pack(anchor='w', pady=(5, 0))

        # **Use tk.Label instead of tk.Entry for display-only path**
        self.service_file_label = tk.Label(
            self.service_text_frame,
            textvariable=self.service_account_file_path_var,
            bg=button_bg,
            fg=fg,
            wraplength=500,
            justify='left',
            font=('TkDefaultFont', 8)
        )
        self.service_file_label.pack(anchor='w', pady=(0, 5))

        # Bind events to all child widgets within service_file_button
        self.bind_events_to_children(
            self.service_file_button,
            self.select_service_account_file,
            self.service_account_file_drop
        )

        # --- Google Sheets Spreadsheet ID ---
        # Fetch the latest Spreadsheet ID from the configuration
        current_spreadsheet_id = self.get_current_spreadsheet_id()
        self.sheets_id_var = tk.StringVar(value=current_spreadsheet_id)

        self.sheets_id_button = tk.Frame(
            dialog,
            bg=button_bg,
            relief='raised',
            bd=2
        )
        self.sheets_id_button.pack(pady=10, fill='x', padx=50)

        self.sheets_id_button.bind("<Enter>", lambda e: self.sheets_id_button.config(relief='groove'))
        self.sheets_id_button.bind("<Leave>", lambda e: self.sheets_id_button.config(relief='raised'))
        self.sheets_id_button.bind("<Button-1>", lambda e: self.sheets_id_entry.focus_set())

        # Place the icon inside the button frame
        self.sheets_id_icon_label = tk.Label(
            self.sheets_id_button,
            image=self.upload_sheets_id_icon,
            bg=button_bg
        )
        self.sheets_id_icon_label.pack(side='left', padx=(20, 5), pady=15)

        # Create a frame for the label and entry widget
        self.sheets_text_frame = tk.Frame(
            self.sheets_id_button,
            bg=button_bg
        )
        self.sheets_text_frame.pack(side='left', pady=10)

        # Place the label and entry widget side by side
        self.sheets_id_text_label = tk.Label(
            self.sheets_text_frame,
            text="Spreadsheet ID:",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.sheets_id_text_label.pack(side='left', padx=(0, 5))

        # **Keep tk.Entry for editable Spreadsheet ID without wraplength**
        self.sheets_id_entry = self.create_entry(
            self.sheets_text_frame,
            textvariable=self.sheets_id_var,
            bg=entry_bg,
            fg=entry_fg,
            font=('TkDefaultFont', 8),
            bd=1,
            highlightthickness=2.5,
            relief='sunken',
            width=45
        )
        self.sheets_id_entry.pack(side='left')

        # Bind events to all child widgets within sheets_id_button
        self.bind_events_to_children(
            self.sheets_id_button,
            lambda e: self.sheets_id_entry.focus_set(),
            None  # Assuming no drag-and-drop for Spreadsheet ID
        )

        # Additionally, bind events directly to the frame to cover any gaps
        self.sheets_id_button.bind("<Button-1>", lambda e: self.sheets_id_entry.focus_set())

        # Save Changes button using ttk.Button with custom style
        ttk.Button(
            dialog,
            text="Save Changes",
            command=lambda: self.save_settings(dialog),
            style="Custom.TButton"
        ).pack(pady=20)

    def open_applications_config_dialog(self):
        """Open a dialog to configure the Applications.xlsx file."""
        dialog = tk.Toplevel(self)
        dialog.title("Applications File Configuration")
        dialog.geometry("600x200")  # Adjusted size for content

        # Keep settings dialog in front
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus()

        # Set dialog theme based on current mode
        bg = self.bg_color
        fg = self.fg_color
        entry_bg = self.entry_bg_color
        entry_fg = self.entry_fg_color
        button_bg = self.button_bg_color

        dialog.config(bg=bg)

        # --- Applications.xlsx file ---
        self.app_file_path_var = tk.StringVar(value=self.get_current_applications_file_path())
        self.app_file_button = tk.Frame(dialog, bg=button_bg, relief='raised', bd=2)
        self.app_file_button.pack(pady=20, fill='x', padx=50)

        # Bind the frame directly to trigger file selection and drag-and-drop
        self.app_file_button.bind("<Enter>", lambda e: self.app_file_button.config(relief='groove'))
        self.app_file_button.bind("<Leave>", lambda e: self.app_file_button.config(relief='raised'))
        self.app_file_button.bind("<Button-1>", lambda e: self.select_app_file())

        # Use tkinterdnd2 methods for drag-and-drop
        self.app_file_button.drop_target_register(DND_FILES)
        self.app_file_button.dnd_bind('<<Drop>>', self.app_file_drop)

        # Place the icon and labels inside the frame
        self.app_icon_label = tk.Label(self.app_file_button, image=self.upload_xlsx_icon, bg=button_bg)
        self.app_icon_label.pack(side='left', padx=(10, 5), pady=10)

        # Create a frame for the text labels
        self.app_text_frame = tk.Frame(self.app_file_button, bg=button_bg)
        self.app_text_frame.pack(side='left', fill='x', expand=True)

        # Text label for instructions
        self.app_text_label = tk.Label(
            self.app_text_frame,
            text="Upload or Drop Applications.xlsx File Here",
            bg=button_bg,
            fg=fg,
            font=('TkDefaultFont', 12)
        )
        self.app_text_label.pack(anchor='w', pady=(5, 0))

        # Path label
        self.app_file_label = tk.Label(
            self.app_text_frame,
            textvariable=self.app_file_path_var,
            bg=button_bg,
            fg=fg,
            wraplength=500,
            justify='left',
            font=('TkDefaultFont', 8)
        )
        self.app_file_label.pack(anchor='w', pady=(0, 5))

        # Bind events to all child widgets within app_file_button
        self.bind_events_to_children(
            self.app_file_button,
            self.select_app_file,
            self.app_file_drop
        )

        # Save Changes button using ttk.Button with custom style
        ttk.Button(
            dialog,
            text="Save Changes",
            command=lambda: self.save_applications_settings(dialog),
            style="Custom.TButton"
        ).pack(pady=20)

    def open_clipboard_editor(self):
        """
        Open a window to edit personal information stored in personal_info.json,
        including masking options.
        """

        # Check if the clipboard editor window already exists
        if hasattr(self, "clipboard_editor_window") and self.clipboard_editor_window.winfo_exists():
            # Bring the existing window to focus
            self.clipboard_editor_window.focus_set()
            return

        # Create a new Toplevel window
        self.clipboard_editor_window = tk.Toplevel(self)
        self.clipboard_editor_window.title("Edit Clipboard")
        self.clipboard_editor_window.geometry("550x525")  # Adjusted size as needed
        self.clipboard_editor_window.transient(self)  # Keep the window on top of the main application
        self.clipboard_editor_window.grab_set()  # Make the window modal

        # Set the window background to match the current theme
        self.clipboard_editor_window.config(bg=self.bg_color)

        # Load personal information
        personal_info = load_personal_info()

        # Store references to entry widgets and masking variables
        entries = {}
        self.mask_vars = {}  # To store BooleanVars for masking

        # Create a frame to hold all widgets with padding
        frame = tk.Frame(self.clipboard_editor_window, bg=self.bg_color)
        frame.pack(padx=20, pady=20, fill='both', expand=True)

        # Iterate over personal_info to create labels, entry fields, and masking checkboxes
        for idx, (label, info_dict) in enumerate(personal_info.items()):
            value = info_dict.get("value", "")
            masked = info_dict.get("masked", False)

            # Create a BooleanVar for each mask state
            mask_var = tk.BooleanVar(value=masked)
            self.mask_vars[label] = mask_var  # Store it for later use

            # Label for the key
            label_widget = ttk.Label(
                frame,
                text=label + ":",
                font=self.label_font,
                background=self.bg_color,
                foreground=self.fg_color
            )
            label_widget.grid(row=idx, column=0, sticky="e", padx=20, pady=5)

            # **Use self.create_entry to create Entry widgets**
            entry = self.create_entry(
                frame,
                font=self.entry_font,
                width=40
            )
            entry.insert(0, value)
            entry.grid(row=idx, column=1, sticky="w", padx=5, pady=5)
            entries[label] = entry

            # Checkbox to toggle masking
            mask_checkbox = tk.Checkbutton(
                frame,
                text="Mask",
                variable=mask_var,
                command=lambda lbl=label: self.toggle_mask(lbl),
                bg=self.bg_color,
                fg=self.fg_color,
                activebackground=self.menu_active_bg,
                activeforeground=self.fg_color,
                selectcolor=self.entry_bg_color
            )
            mask_checkbox.grid(row=idx, column=2, sticky="w", padx=5, pady=5)

        # Save and Cancel buttons
        button_frame = tk.Frame(frame, bg=self.bg_color)
        button_frame.grid(row=len(personal_info), column=0, columnspan=3, pady=20)

        save_button = ttk.Button(
            button_frame,
            text="Save",
            command=lambda: self.save_clipboard_info(entries, self.clipboard_editor_window),
            style="Custom.TButton"
        )
        save_button.pack(side="left", padx=10)

        cancel_button = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.clipboard_editor_window.destroy,
            style="Custom.TButton"
        )
        cancel_button.pack(side="left", padx=10)

    def save_settings(self, dialog):
        """Save settings related to Google Sync and close the dialog."""
        # Retrieve values from the UI
        service_account_path = self.service_account_file_path_var.get().strip()
        spreadsheet_id = self.sheets_id_var.get().strip()

        # Validate inputs
        if not os.path.isfile(service_account_path):
            messagebox.showerror("Error", "Service Account JSON file does not exist.")
            return
        if not spreadsheet_id:
            messagebox.showerror("Error", "Spreadsheet ID cannot be empty.")
            return

        # Update configuration
        try:
            # Load existing configuration
            with open(CONFIG_JSON_PATH, "r") as config_file:
                config_data = json.load(config_file)

            # Update the relevant fields
            config_data["SERVICE_ACCOUNT_FILE"] = service_account_path
            config_data["SPREADSHEET_ID"] = spreadsheet_id

            # Save updated configuration
            with open(CONFIG_JSON_PATH, "w") as config_file:
                json.dump(config_data, config_file, indent=4)

            print("[DEBUG] Settings saved successfully.")
            messagebox.showinfo("Success", "Your Google Sync settings have been successfully saved. Please restart the application to apply the changes.")

            # Optionally, reload configurations if needed
            self.reload_configurations()

            dialog.destroy()
        except Exception as e:
            print(f"Error: Failed to save settings: {e}")
            logging.error(f"Error: Failed to save settings: {e}")
            messagebox.showerror("Error", f"Failed to save settings: {e}")

    def save_applications_settings(self, dialog):
        """Save settings related to Applications.xlsx and close the dialog."""
        # Retrieve values from the UI
        applications_path = self.app_file_path_var.get()

        # Validate inputs
        if not os.path.isfile(applications_path):
            messagebox.showerror("Error", "Applications.xlsx file does not exist.")
            return

        # Update configuration (Assuming you have a method or mechanism to handle configurations)
        try:
            # Example: Update a configuration dictionary or write to a config file
            config = {
                "DATA_FILE_PATH": applications_path
            }
            with open('config.json', 'w') as config_file:
                json.dump(config, config_file, indent=4)
            print("[DEBUG] Applications settings saved successfully.")
            messagebox.showinfo("Success", "Your Application settings have been successfully saved. Please restart the application to apply the changes.")
            dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Applications settings: {e}")

    def save_clipboard_info(self, entries, window):
        """
        Save the edited personal information to personal_info.json and refresh the Personal Information tab.

        Parameters:
        - entries (dict): A dictionary mapping keys to their respective Entry widgets.
        - window (tk.Toplevel): The editing window to be closed after saving.
        """
        # Construct a dictionary with updated values and masked flags
        updated_info = {}
        for key, entry in entries.items():
            masked = self.mask_vars.get(key, tk.BooleanVar()).get()
            updated_info[key] = {
                "value": entry.get().strip(),
                "masked": masked
            }

        # Validate required fields if necessary
        if not updated_info.get("Email", {}).get("value"):
            messagebox.showerror("Validation Error", "Email cannot be empty.")
            return

        # Save the updated information to the JSON file
        try:
            with open(PERSONAL_INFO_FILE, "w") as file:
                json.dump(updated_info, file, indent=4)
            print("Personal information updated successfully.")
        except Exception as e:
            print(f"Error saving personal information: {e}")
            messagebox.showerror("Error", f"Failed to save personal information: {e}")
            return

        # Refresh the Personal Information tab to reflect changes
        self.create_personal_info_tab()

        # Close the editing window
        window.destroy()

        # Notify the user of successful save
        messagebox.showinfo("Success", "Personal information updated successfully.")

    def get_current_google_sync_setting(self):
        """Retrieve the current ENABLE_GOOGLE_SYNC setting from settings_manager.py."""
        try:
            return ENABLE_GOOGLE_SYNC
        except ImportError:
            return False  # Default to False if not set

    # Additional methods for file selection, Data handling, and layout setup

    def get_current_applications_file_path(self):
        """Retrieve the current Applications.xlsx file path from settings_manager.py."""
        try:
            from config.settings_manager import DATA_FILE_PATH
            if os.path.isfile(DATA_FILE_PATH):
                return os.path.abspath(DATA_FILE_PATH)
            else:
                print("DATA_FILE_PATH does not point to an existing file.")
                return "No file selected"
        except Exception as e:
            print(f"Error retrieving DATA_FILE_PATH: {e}")
            return "No file selected"

    def get_current_spreadsheet_id(self):
        """Retrieve the current Spreadsheet ID from config.settings_manager."""
        try:
            from config.settings_manager import SPREADSHEET_ID
            return SPREADSHEET_ID
        except ImportError:
            print("Error: Could not import SPREADSHEET_ID from config.settings_manager.")
            logging.error("Error: Could not import SPREADSHEET_ID from config.settings_manager.")
            return ""
        except Exception as e:
            print(f"Error retrieving SPREADSHEET_ID: {e}")
            logging.error(f"Error retrieving SPREADSHEET_ID: {e}")
            return ""

    def get_current_service_account_file_path(self):
        """Retrieve the current Service Account JSON file path from config.settings_manager."""
        try:
            from config.settings_manager import SERVICE_ACCOUNT_FILE
            if os.path.isfile(SERVICE_ACCOUNT_FILE):
                return os.path.abspath(SERVICE_ACCOUNT_FILE)
            else:
                print("SERVICE_ACCOUNT_FILE does not point to an existing file.")
                return "No file selected"
        except ImportError:
            return "No file selected"
        except Exception as e:
            print(f"Error retrieving SERVICE_ACCOUNT_FILE: {e}")
            logging.error(f"Error retrieving SERVICE_ACCOUNT_FILE: {e}")
            return "No file selected"

    def update_config(self, **kwargs):
        """Update configuration settings in app_config.json."""
        try:
            # Load the current configuration
            if os.path.exists(self.CONFIG_JSON_PATH):
                with open(self.CONFIG_JSON_PATH, "r") as config_file:
                    config = json.load(config_file)
            else:
                # Start with default configuration if file doesn't exist
                config = default_config.copy()
                print("[DEBUG] Config file not found. Using default configuration.")

            # Update the configuration with new values
            config.update(kwargs)

            # Save the updated configuration back to app_config.json
            with open(self.CONFIG_JSON_PATH, "w") as config_file:
                json.dump(config, config_file, indent=4)
            print(f"[DEBUG] Configuration updated: {kwargs}")
        except Exception as e:
            print(f"Error updating configuration: {e}")
            logging.error(f"Error updating configuration: {e}")

    def reload_configurations(self):
        """Reload configurations from app_config.json."""
        try:
            with open(self.CONFIG_JSON_PATH, "r") as config_file:
                config = json.load(config_file)

            # Update variables
            self.sync_to_google = config.get("ENABLE_GOOGLE_SYNC", False)
            self.DATA_FILE_PATH = config.get("DATA_FILE_PATH", os.path.join(base_path, "Data", "Applications.xlsx"))
            self.SERVICE_ACCOUNT_FILE = config.get("SERVICE_ACCOUNT_FILE",
                                                   os.path.join(base_path, "config", "service_account.json"))
            self.SPREADSHEET_ID = config.get("SPREADSHEET_ID", "")
            theme = config.get("theme", "Light")
            self.clipboard_side = config.get("clipboard_side", "right")  # Load clipboard_side
            self.clipboard_enabled = config.get("clipboard_enabled", True)  # Load clipboard_enabled

            print(
                f"[DEBUG] Reloading configurations. Theme: {theme}, Clipboard Side: {self.clipboard_side}, Clipboard Enabled: {self.clipboard_enabled}")

            # Update theme
            self.is_dark_mode = True if theme.lower() == "dark" else False
            self.apply_theme()

            # Reconfigure the main layout with the updated clipboard_side
            self.setup_main_layout()

            # Re-read the Excel file with the updated path
            try:
                self.applications_df = read_applications_from_excel(self.DATA_FILE_PATH)
                self.populate_treeview(self.applications_df)
                print("[DEBUG] Applications Data reloaded successfully.")
            except Exception as e:
                print(f"[ERROR] Could not read the Excel file after reloading configurations: {e}")
                logging.error(f"[ERROR] Could not read the Excel file after reloading configurations: {e}")
                self.applications_df = pd.DataFrame()
                self.populate_treeview(self.applications_df)

            # Re-establish Google Sync if enabled
            if self.sync_to_google:
                self.sync_to_google_sheets()
                self.schedule_sync()

            print("[DEBUG] Configurations reloaded successfully.")
        except Exception as e:
            print(f"[ERROR] Failed to reload configurations: {e}")
            logging.error(f"[ERROR] Failed to reload configurations: {e}")
            messagebox.showerror("Error", f"Failed to reload configurations: {e}")

    def update_google_sync_setting(self, enable_google_sync):
        """
        Update ENABLE_GOOGLE_SYNC setting in app_config.json.
        """
        try:
            # Load current config
            with open(self.CONFIG_JSON_PATH, "r") as file:
                config = json.load(file)

            # Update ENABLE_GOOGLE_SYNC value
            config["ENABLE_GOOGLE_SYNC"] = enable_google_sync

            # Write updated config back to file
            with open(self.CONFIG_JSON_PATH, "w") as file:
                json.dump(config, file, indent=4)

            print(f"Google Sync setting updated to: {enable_google_sync}")
        except Exception as e:
            print(f"Error updating Google Sync setting: {e}")


if __name__ == "__main__":
    app = AppTrack()
    app.mainloop()
