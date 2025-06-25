# core.py
import json
import hashlib
import importlib.util
import os
import ctypes
import shutil
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import win32gui
import win32con
from datetime import datetime
import requests

# Constants for window styles
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Update configuration
GITHUB_REPO = "https://api.github.com/repos/KyleJamesOlson/RefurbHelper/contents/RWH"
LOCAL_CACHE = os.path.join(os.path.dirname(sys.executable), "module_cache")
VERSIONS_FILE = "versions.json"
CURRENT_VERSIONS = {
    "parser.py": "1.0.0",
    "template.py": "1.0.0",
    "utils.py": "1.0.0"
}

def calculate_sha256(file_path):
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        while chunk := f.read(8192):
            sha256.update(chunk)
    return sha256.hexdigest()

def load_module(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None:
        logging.error(f"Failed to create spec for {file_path}")
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def check_for_updates():
    try:
        os.makedirs(LOCAL_CACHE, exist_ok=True)
        api_url = GITHUB_API_URL
        logging.debug(f"Checking updates from API: {api_url}")
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()  # Raise an exception for bad status codes (e.g., 404)
        contents = response.json()
        logging.debug(f"API response: {contents}")

        # Fetch versions.json to get version and SHA information
        versions_url = next(item["download_url"] for item in contents if item["name"] == VERSIONS_FILE)
        versions_response = requests.get(versions_url, timeout=10)
        versions_response.raise_for_status()
        server_versions = versions_response.json().get("modules", {})
        logging.debug(f"Loaded versions: {server_versions}")

        for item in contents:
            module_name = item["name"]
            if module_name not in ["parser.py", "template.py", "utils.py"]:
                continue
            server_info = server_versions.get(module_name, {})
            server_version = server_info.get("version", "0.0.0")
            server_sha256 = server_info.get("sha256", "")
            server_url = item["download_url"]
            local_file = os.path.join(LOCAL_CACHE, module_name)
            current_version = CURRENT_VERSIONS.get(module_name, "0.0.0")
            logging.debug(f"Checking {module_name}: Current={current_version}, Server={server_version}")
            if server_version > current_version:
                logging.info(f"Update available for {module_name}: {current_version} -> {server_version}")
                response = requests.get(server_url, timeout=10)
                response.raise_for_status()
                with open(local_file, 'wb') as f:
                    f.write(response.content)
                local_sha256 = calculate_sha256(local_file)
                if local_sha256 != server_sha256:
                    logging.error(f"Checksum mismatch for {module_name}. Discarding update.")
                    os.remove(local_file)
                    continue
                CURRENT_VERSIONS[module_name] = server_version
                logging.info(f"Successfully downloaded {module_name} v{server_version}")
        for module_name in server_versions:
            local_file = os.path.join(LOCAL_CACHE, module_name)
            if os.path.exists(local_file):
                module = load_module(module_name.replace('.py', ''), local_file)
                if module:
                    logging.info(f"Loaded updated module {module_name}")
                    if module_name == "parser.py":
                        globals()['parse_txt_file'] = module.parse_txt_file
                    elif module_name == "template.py":
                        globals()['fill_template'] = module.fill_template
                    elif module_name == "utils.py":
                        globals()['replace_in_runs'] = module.replace_in_runs
                        globals()['process_element'] = module.process_element
    except requests.exceptions.RequestException as e:
        logging.error(f"Update check failed (network issue): {str(e)}", exc_info=True)
        messagebox.showwarning("Update Warning", "Failed to check for updates from GitHub. Using default modules.")
    except Exception as e:
        logging.error(f"Update check failed: {str(e)}", exc_info=True)
        messagebox.showwarning("Update Warning", "Failed to check for updates. Using default modules.")


# Fallback to bundled modules
if getattr(sys, 'frozen', False):
    for module_name in ["parser.py", "template.py", "utils.py"]:
        if module_name == "parser.py" and 'parse_txt_file' not in globals():
            bundled_file = os.path.join(sys._MEIPASS, 'assets', module_name)
            module = load_module(module_name.replace('.py', ''), bundled_file)
            globals()['parse_txt_file'] = module.parse_txt_file
        elif module_name == "template.py" and 'fill_template' not in globals():
            bundled_file = os.path.join(sys._MEIPASS, 'assets', module_name)
            module = load_module(module_name.replace('.py', ''), bundled_file)
            globals()['fill_template'] = module.fill_template
        elif module_name == "utils.py" and 'process_element' not in globals():
            bundled_file = os.path.join(sys._MEIPASS, 'assets', module_name)
            module = load_module(module_name.replace('.py', ''), bundled_file)
            globals()['replace_in_runs'] = module.replace_in_runs
            globals()['process_element'] = module.process_element
else:
    from parser import parse_txt_file
    from template import fill_template
    from utils import replace_in_runs, process_element

check_for_updates()

class AssetFormFiller(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Spectrum E-cycle Refurb Worksheet Helper")
        self.geometry("540x600")
        self.resizable(False, False)
        self.overrideredirect(False)
	#self.attributes('-topmost', True)

        # Set the window icon
        try:
            if getattr(sys, 'frozen', False):  # Running as a PyInstaller bundle
                icon_path = os.path.join(sys._MEIPASS, 'assets', 'icon.ico')
            else:  # Running in development
                icon_path = os.path.join("assets", "icon.ico")
            self.iconbitmap(default=icon_path)
        except Exception as e:
            logging.error(f"Failed to set window icon: {str(e)}")

        # Adjust window style to show in taskbar
        self.set_appwindow()

        self.title_bar = tk.Frame(self, bg="#559546", height=30)
        self.title_bar.pack(fill=tk.X)
        
        self.title_label = tk.Label(
            self.title_bar, 
            text="Spectrum E-cycle Refurb Worksheet Helper", 
            bg="#559546", 
            fg="white", 
            font=("Roboto", 12, "bold"),
            padx=10
        )
        self.title_label.pack(side=tk.LEFT, pady=5)

        try:
            close_img = Image.open(os.path.join("assets", "icon.ico")).resize((20, 20), Image.LANCZOS)
            self.close_icon = ImageTk.PhotoImage(close_img)
            self.close_button = tk.Button(
                self.title_bar, 
                image=self.close_icon, 
                bg="#f4f3f1", 
                borderwidth=0, 
                command=self.quit
            )
            self.close_button.pack(side=tk.RIGHT, padx=5)
        except Exception as e:
            logging.error(f"Failed to load close button icon: {str(e)}")
            self.close_button = tk.Button(
                self.title_bar, 
                text="X", 
                bg="#559546", 
                fg="white", 
                font=("Roboto", 12, "bold"),
                borderwidth=0, 
                command=self.quit
            )
            self.close_button.pack(side=tk.RIGHT, padx=5)

        self.minimize_button = tk.Button(
            self.title_bar, 
            text="_", 
            bg="#559546", 
            fg="white", 
            font=("Roboto", 12, "bold"),
            borderwidth=0, 
            command=self.minimize_window
        )
        self.minimize_button.pack(side=tk.RIGHT, padx=5)

        self.title_bar.bind("<Button-1>", self.start_move)
        self.title_bar.bind("<B1-Motion>", self.on_motion)
        self.title_label.bind("<Button-1>", self.start_move)
        self.title_label.bind("<B1-Motion>", self.on_motion)

        self.main_frame = tk.Frame(self, bg="#f0f0f0", bd=2, relief="flat")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        self.main_frame.pack_propagate(False)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "TButton", 
            font=("Roboto", 10, "bold"), 
            background="#559546", 
            foreground="black", 
            bordercolor="#559546", 
            borderwidth=1, 
            padding=6
        )
        style.map(
            "TButton", 
            background=[('active', '#46803b')],
            foreground=[('active', 'white')]
        )
        style.configure(
            "TEntry", 
            font=("Roboto", 10), 
            fieldbackground="white", 
            foreground="black", 
            bordercolor="#cccccc", 
            borderwidth=1, 
            padding=6
        )
        style.map(
            "TEntry", 
            fieldbackground=[('focus', '#e6f3fa')],
            bordercolor=[('focus', '#559546')]
        )
        style.configure(
            "TCheckbutton", 
            font=("Roboto", 10, "bold"), 
            background="#f0f0f0"
	)
        try:
            bg_img = Image.open(os.path.join("assets", "background.png")).resize((540, 600), Image.LANCZOS)
            self.background_image = ImageTk.PhotoImage(bg_img)
            self.background_label = tk.Label(self, image=self.background_image)
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
            self.background_label.lower()
        except Exception as e:
            logging.error(f"Failed to load background image: {str(e)}")
            self.configure(bg="#f0f0f0")

        try:
            logo_img = Image.open(os.path.join("assets", "Logo1.png"))
            img_width, img_height = logo_img.size
            target_width = 540
            target_height = int((target_width / img_width) * img_height)
            if target_height > 100:
                target_height = 100
                target_width = int((target_height / img_height) * img_width)
            logo_img = logo_img.resize((target_width, target_height), Image.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(logo_img)
            self.logo_label = tk.Label(self.main_frame, image=self.logo_image, bg="#f0f0f0")
            self.logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 0), sticky="ew")
        except Exception as e:
            logging.error(f"Failed to load logo image: {str(e)}")
            self.logo_label = tk.Label(self.main_frame, text="Spectrum E-cycle", font=("Roboto", 14, "bold"), bg="#f0f0f0")
            self.logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 0), sticky="ew")

        self.data_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.technician_initials = tk.StringVar()
        self.warranty = tk.BooleanVar()
        self.warranty_date = tk.StringVar()
        self.power_adaptor = tk.BooleanVar()
        self.touchscreen = tk.BooleanVar()
        self.ports = tk.StringVar()
        self.condition = tk.StringVar()

        default_template = os.path.join(os.getcwd(), "Template.docx")
        if os.path.exists(default_template):
            self.template_path.set(default_template)
        else:
            logging.warning(f"Default template file {default_template} not found")

        self.create_file_inputs()
        self.create_form_inputs()
        self.create_submit_button()

    def set_appwindow(self):
        try:
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = style & ~WS_EX_TOOLWINDOW
            style = style | WS_EX_APPWINDOW
            res = ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            self.wm_withdraw()
            self.after(10, lambda: self.wm_deiconify())
        except Exception as e:
            logging.error(f"Failed to set appwindow style: {str(e)}")

    def start_move(self, event):
        self.x_root = event.x_root
        self.y_root = event.y_root

    def on_motion(self, event):
        deltax = event.x_root - self.x_root
        deltay = event.y_root - self.y_root
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry(f"+{x}+{y}")
        self.x_root = event.x_root
        self.y_root = event.y_root

    def minimize_window(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception as e:
            logging.error(f"Failed to minimize window: {str(e)}")
            messagebox.showerror("Error", "Unable to minimize window")

    def create_file_inputs(self):
        tk.Label(self.main_frame, text="HWINFO Log:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.data_path, width=40, style="TEntry").grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=lambda: self.browse_file(self.data_path, [("Text and Log files", "*.txt *.log")]), style="TButton").grid(row=1, column=2, padx=10, pady=5)

        tk.Label(self.main_frame, text="Template File:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.template_path, width=40, style="TEntry").grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=lambda: self.browse_file(self.template_path, [("Word documents", "*.docx")]), style="TButton").grid(row=2, column=2, padx=10, pady=5)

        tk.Label(self.main_frame, text="Output Directory:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.output_path, width=40, style="TEntry").grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(self.main_frame, text="Browse", command=lambda: self.browse_directory(self.output_path), style="TButton").grid(row=3, column=2, padx=10, pady=5)

    def create_form_inputs(self):
        tk.Label(self.main_frame, text="Technician Initials:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=4, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.technician_initials, width=40, style="TEntry").grid(row=4, column=1, padx=5, pady=5)

        ttk.Checkbutton(self.main_frame, text="Warranty", variable=self.warranty, command=self.toggle_warranty_date, style="TCheckbutton").grid(row=5, column=0, padx=10, pady=5, sticky="e")
        self.warranty_date_entry = ttk.Entry(self.main_frame, textvariable=self.warranty_date, width=40, state='disabled', style="TEntry")
        self.warranty_date_entry.grid(row=5, column=1, padx=5, pady=5)

        ttk.Checkbutton(self.main_frame, text="Power Adaptor", variable=self.power_adaptor, style="TCheckbutton").grid(row=6, column=0, padx=10, pady=5, sticky="e")
        ttk.Checkbutton(self.main_frame, text="Touchscreen?", variable=self.touchscreen, style="TCheckbutton").grid(row=7, column=0, padx=10, pady=5, sticky="e")

        tk.Label(self.main_frame, text="Ports:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=8, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.ports, width=40, style="TEntry").grid(row=8, column=1, padx=5, pady=5)

        tk.Label(self.main_frame, text="Condition:", bg="#f0f0f0", fg="black", font=("Roboto", 10, "bold")).grid(row=9, column=0, padx=10, pady=5, sticky="e")
        ttk.Entry(self.main_frame, textvariable=self.condition, width=40, style="TEntry").grid(row=9, column=1, padx=5, pady=5)

    def browse_file(self, path_var, file_types):
        initial_dir = os.path.join(os.getcwd(), "Logs") if path_var == self.data_path else os.getcwd()
        if not os.path.exists(initial_dir):
            os.makedirs(initial_dir)
        file_path = filedialog.askopenfilename(filetypes=file_types, initialdir=initial_dir)
        if file_path:
            path_var.set(file_path)
            if path_var == self.data_path:
                self.output_path.set(os.path.dirname(file_path))

    def browse_directory(self, path_var):
        initial_dir = os.path.dirname(self.data_path.get()) if self.data_path.get() else os.getcwd()
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
        dir_path = filedialog.askdirectory(initialdir=initial_dir)
        if dir_path:
            path_var.set(dir_path)

    def toggle_warranty_date(self):
        state = 'normal' if self.warranty.get() else 'disabled'
        self.warranty_date_entry.config(state=state)

    def submit(self):
        if not all([self.data_path.get(), self.template_path.get(), self.output_path.get(),
                   self.technician_initials.get(), self.ports.get(), self.condition.get()]):
            messagebox.showerror("Error", "Please fill all required fields")
            return

        if self.warranty.get() and not self.warranty_date.get():
            messagebox.showerror("Error", "Please enter warranty expiration date")
            return

        try:
            data, camera_found, keyname_fallback = parse_txt_file(self.data_path.get())
            brand_name = data.get(next((key for key in data if 'Computer Brand Name' in data[key]), None), {}).get('Computer Brand Name', 'Unknown').replace(" ", "_")
            serial_number = data.get(next((key for key in data if 'Product Serial Number' in data[key]), None) or next((key for key in data if key == "System"), None), {}).get('Product Serial Number', 'Unknown')
            current_date = datetime.now().strftime("%m/%d/%Y").replace("/", "")
            output_file = os.path.join(self.output_path.get(), f"{brand_name}_{serial_number}_{current_date}.docx")
            form_data = {
                'technician_initials': self.technician_initials.get(),
                'warranty': self.warranty.get(),
                'warranty_date': self.warranty_date.get(),
                'power_adaptor': self.power_adaptor.get(),
                'touchscreen': self.touchscreen.get(),
                'ports': self.ports.get(),
                'condition': self.condition.get()
            }
            fill_template(self.template_path.get(), output_file, data, camera_found, form_data, keyname_fallback)
            messagebox.showinfo("Success", f"Form filled and saved as {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def create_submit_button(self):
        ttk.Button(self.main_frame, text="Generate Form", command=self.submit, style="TButton").grid(row=10, column=1, pady=20)

if __name__ == "__main__":
    try:
        app = AssetFormFiller()
        logging.debug("AssetFormFiller initialized successfully")
        app.mainloop()
    except Exception as e:
        logging.error(f"Failed to initialize AssetFormFiller: {str(e)}")
        raise
