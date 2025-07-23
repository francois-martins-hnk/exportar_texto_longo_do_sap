# %%
# - - - - - Libraries and Imports
import configparser
import psutil
import pandas as pd
from pywinauto import Application
import time
import pyperclip  
from math import e
import warnings
warnings.filterwarnings("ignore")
import os
import configparser

# %%
# - - - - - Parameters

# Get the absolute path to the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
# print(f"Script directory: {script_dir}")

# Build the full path to config.ini in the same folder
config_path = os.path.join(script_dir, 'config.ini')
print(f"Config file path: {config_path}")

# Load config
config = configparser.ConfigParser()
config.read(config_path)
conf = config['DEFAULT']
time_out = int(config['DEFAULT']['time_out'])
sap_logon_path = config['DEFAULT']['sap_logon_path']
file_path = config['DEFAULT']['file_path']

# %%
# - - - - - Functions

# This function retrieves the long text of a material from SAP using the MMBE transaction.
def get_material_long_text(material_code, dlg):
  status = 1 # 1 for success, 0 for failure
  try: 
    # Connect to SAP Easy Access Without Password
    dlg = app.window(title_re="SAP Easy.*")
    dlg.wait('ready', timeout=time_out)

    # Navigate to the MMBE transaction
    dlg.type_keys('MMBE')
    time.sleep(time_out/10)
    dlg.type_keys('{ENTER}')
    dlg.wait('ready', timeout=time_out)
    time.sleep(time_out/10)
    dlg = app.window(title_re=".*Visão.*")
    dlg.wait('ready', timeout=time_out)
    time.sleep(time_out/10)

    # Type the material code and navigate through the dialog
    dlg.type_keys(material_code)
    dlg.wait('ready', timeout=time_out)
    dlg.type_keys('{F8}')
    time.sleep(time_out/2)
    
    # Navigate to the long text section (MM03)
    dlg.type_keys('{F9}')
    time.sleep(time_out/10)
    dlg = app.window(title_re=".*Exibir.*")
    dlg.wait('ready', timeout=time_out)
    time.sleep(time_out/10)
    dlg.type_keys('+{TAB}')
    for _ in range(10):
      dlg.type_keys('{RIGHT}')
      time.sleep(time_out/20)
    dlg.type_keys('{ENTER}') 
    dlg.wait('ready', timeout=time_out)

    # Navigate to the text area
    for _ in range(10):
      dlg.type_keys('{TAB}')
      time.sleep(time_out/20)

    # Select all text and copy it
    dlg.type_keys('^a') 
    time.sleep(time_out/10)
    dlg.type_keys('^c') 
    time.sleep(time_out/10)

    # Retrieve the copied text from the clipboard
    material_long_text = pyperclip.paste()

    # Close the dialog and return to the SAP Easy Access screen
    dlg.type_keys('{F3}')
    dlg.wait('ready', timeout=time_out)
    dlg = app.window(title_re=".*Visão.*")
    dlg.wait('ready', timeout=time_out)
    dlg.type_keys('{F12}')
    time.sleep(time_out/10)
    dlg.wait('ready', timeout=time_out)
    dlg.type_keys('{F12}')
    time.sleep(time_out/10)
    dlg = app.window(title_re="SAP Easy.*")
    
    status = True
    return status, material_long_text
  
  except Exception as e:
    status = 0
    material_long_text = ""
    return status, material_long_text


# %%
# - - - - - Main Script
# This script processes a list of materials from an Excel file, retrieves their long descriptions from SAP, and updates the Excel file with the retrieved descriptions.

# Load the Excel file containing material codes
try:
  df = pd.read_excel(file_path)
  total_rows = len(df)
except Exception as e:
  print(f"⚠️ Error reading the Excel file: {e}")
  input("\n✅ Script completed. Press Enter to exit...")
  exit(1)
  raise SystemExit(1)
  
# Check and ensure required columns
required_columns = ["material_code"]
missing_cols = [col for col in required_columns if col not in df.columns]

if missing_cols:
    print(f"⚠️ Missing required columns in the Excel file: {missing_cols}")
    input("\n✅ Script completed. Press Enter to exit...")
    exit(1)
    raise SystemExit(1)

# Create the 'long_description' column if it doesn't exist
if "long_description" not in df.columns:
    df["long_description"] = pd.NA

# Check if there are any material codes to process
if df["material_code"].isnull().all():
    print("⚠️ No material codes found in the Excel file.")
    input("\n✅ Script completed. Press Enter to exit...")
    exit(1)
    raise SystemExit(1)


# Identify rows to process: expanded_text is blank AND both descriptions are present
to_process_mask = (
    df["long_description"].isnull() | (df["long_description"].astype(str).str.strip() == "")
) & df["material_code"].notnull()

rows_to_process = df[to_process_mask]

# If nothing to process
if rows_to_process.empty:
  print("[ ] All rows already have long_description filled or are missing descriptions.")
else:
  
  # --- Kill any existing SAP Logon instances ---
  for proc in psutil.process_iter(attrs=['pid', 'name']):
    if proc.info['name'] and "saplogon.exe" in proc.info['name'].lower():
      try:
        proc.kill()
        # print("🔁 Existing SAP Logon process killed.")
      except psutil.NoSuchProcess:
        pass
      except Exception as e:
        print(f"⚠️ Could not kill process: {e}")

  # --- Start SAP Logon ---
  app = Application(backend="win32").start(sap_logon_path)
  dlg = app.window(title_re=".*SAP Logon.*")
  dlg.wait('ready', timeout=time_out)
  time.sleep(time_out / 10)
  dlg.type_keys('{ENTER}')
  time.sleep(time_out)
  
  # print(f"[ ] Found {len(rows_to_process)} materials to process out of {total_rows} total materials.\n")
  success = 0
  errors = 0
  start_time = time.time()
  for i, (idx, row) in enumerate(rows_to_process.iterrows(), start=1):
    material_code = row["material_code"]
    
    # For ETA
    elapsed_time = time.time() - start_time  # seconds since start
    avg_time_per_item = elapsed_time / i if i > 0 else 0
    remaining_items = len(rows_to_process) + 1 - i
    eta_seconds = int(avg_time_per_item * remaining_items)
    # Format ETA as H:MM:SS
    eta_str = time.strftime('%H:%M:%S', time.gmtime(eta_seconds))
    

    print(f'🔄 Processing {i}/{len(rows_to_process)}: Material {material_code}   |   '
          f'Successful: {success}   |   '
            f'Errors: {errors}   |   ETA: {eta_str}   ', end="\r")
    # for i, material_code in enumerate(material_list, start=1):
    status, long_text = get_material_long_text(material_code, dlg)
    if status:
      df.at[idx, "long_description"] = long_text
      success += 1
      print(f'✅ Processed {i}/{len(rows_to_process)}: Material {material_code}   |   '
            f'Successful: {success}   |   '
            f'Errors: {errors}   |   ETA: {eta_str}      ', end="\n")
    else:
      df.at[idx, "long_description"] = ""
      errors += 1
      print(f'❌ Processed {i}/{len(rows_to_process)}: Material {material_code}   |   '
            f'Successful: {success}   |   '
            f'Errors: {errors}   |   ETA: {eta_str}      ', end="\n")
      # --- Kill any existing SAP Logon instances ---
      for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] and "saplogon.exe" in proc.info['name'].lower():
          try:
            proc.kill()
            # print("🔁 Existing SAP Logon process killed.")
          except psutil.NoSuchProcess:
            pass
          except Exception as e:
            print(f"⚠️ Could not kill process: {e}")

      # --- Start SAP Logon ---
      app = Application(backend="win32").start(sap_logon_path)
      dlg = app.window(title_re=".*SAP Logon.*")
      dlg.wait('ready', timeout=time_out)
      time.sleep(time_out / 10)
      dlg.type_keys('{ENTER}')
      time.sleep(time_out)
  
    if i % 5 == 0 or i == len(rows_to_process):
      try:
        df.to_excel(file_path, index=False)
        # print(f"💾 Partial save successful ({i}/{len(rows_to_process)})")
      except Exception as e:
        print(f"⚠️ Error saving the Excel file at iteration {i}: {e}")
        break
  print(f"\n💾 File updated: {file_path} — All changes saved successfully.")
  
  try:
    app.kill()
  except Exception as e:
    print(f"⚠️ Error closing SAP application: {e}")
    
  print(f"\nSummary:")
  print(f"\t[ ] Total materials processed: {len(rows_to_process)}")
  print(f"\t[ ] Successful: {success}")
  print(f"\t[ ] Errors: {errors}")
  print(f'\t[ ] Success[%]: {success / len(rows_to_process) * 100:.2f}%')
  print(f"\t[ ] Total time taken: {time.strftime('%H:%M:%S', time.gmtime(time.time() - start_time))}")
  print(f"\t[ ] File saved at: {file_path}\n")
  
input("\n✅ Script completed. Press Enter to exit...")


