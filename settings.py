import os
import tkinter as tk
from tkinter import filedialog, messagebox
import configparser


# Get the absolute path to the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
# print(f"Script directory: {script_dir}")

# Build the full path to config.ini in the same folder
config_path = os.path.join(script_dir, 'config.ini')

# Load config
config = configparser.ConfigParser()
config.read(config_path)
conf = config['DEFAULT']

def save_config():
    config['DEFAULT'] = {
        'time_out': timeout_var.get(),
        'sap_logon_path': sap_path_var.get(),
        'file_path': file_path_var.get()
    }
    with open(config_path, 'w') as f:
        config.write(f)
    messagebox.showinfo("Saved", "Configuration saved successfully!")

def browse_path(entry_var):
    path = filedialog.askopenfilename()
    if path:
        entry_var.set(path)

# Styling functions for hover effect on buttons
def on_enter(e):
    e.widget['background'] = '#d9d9d9'

def on_leave(e):
    e.widget['background'] = '#f0f0f0'

root = tk.Tk()
root.title("Edit Configuration")
root.geometry("600x180")
root.configure(bg="#f9f9f9")

label_font = ("Segoe UI", 10)
entry_font = ("Segoe UI", 10)

def add_field(row, label_text, variable, browse=False):
    tk.Label(root, text=label_text, bg="#f9f9f9", font=label_font, anchor="w").grid(row=row, column=0, sticky="w", padx=10, pady=6)
    entry = tk.Entry(root, textvariable=variable, font=entry_font, width=54, relief="solid", borderwidth=1)
    entry.grid(row=row, column=1, sticky="w", padx=(0,5), pady=6)
    if browse:
        btn = tk.Button(root, text="Browse", command=lambda: browse_path(variable), bg="#f0f0f0", relief="flat", borderwidth=1,
                        padx=8, pady=3, font=("Segoe UI", 9))
        btn.grid(row=row, column=2, padx=(0,10), pady=6, sticky="w")
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
    return entry

timeout_var = tk.StringVar(value=conf.get('time_out', '15'))
sap_path_var = tk.StringVar(value=conf.get('sap_logon_path', ''))
file_path_var = tk.StringVar(value=conf.get('file_path', ''))

add_field(0, "Timeout (s):", timeout_var)
add_field(1, "SAP Logon Path:", sap_path_var, browse=True)
add_field(2, "Excel File Path:", file_path_var, browse=True)

save_btn = tk.Button(root, text="Save", command=save_config,
                     bg="#f0f0f0", fg="#333", font=("Segoe UI", 10, "bold"),
                     padx=14, pady=6, relief="flat", borderwidth=1)
save_btn.grid(row=4, column=1, pady=12, sticky="w")
save_btn.bind("<Enter>", on_enter)
save_btn.bind("<Leave>", on_leave)

root.mainloop()
