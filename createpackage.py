import pyautogui
import pygetwindow as gw
import pandas as pd
import time
import os

# global delay
pyautogui.PAUSE = 1.5

# Load project names and paths from CSV
csv_path = r'E:\Jsons\Home\PDA\DK\foetex\FolderNames.csv'
try:
    df = pd.read_csv(csv_path)
    print(f"CSV Columns: {df.columns}")
except FileNotFoundError:
    print(f"Error: CSV file not found at {csv_path}!")
    exit()

# Ensure the correct application window is active
windows = gw.getWindowsWithTitle("MobiControl")
if not windows:
    print("MobiControl application not found!")
    exit()

target_window = None
for window in windows:
    if "Package Studio" in window.title:
        target_window = window
        target_window.activate()
        break
if not target_window:
    print("Target application window not found!")
    exit()

# Function to create a single package
def create_package(project_name, json_file_path):
    try:
        print(f"Creating new package: {project_name}")
        target_window.activate()
        time.sleep(1)

        # Start a new package creation
        print(f"starting new package")
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(2)

        # Input project name
        print(f"inputing project name")
        pyautogui.write(project_name)
        pyautogui.press('tab')  # Move to Project Location field
        time.sleep(2)
        pyautogui.press('tab')  # Skip Project Location field
        time.sleep(2)
        pyautogui.press('tab')  # Move to Processor field
        time.sleep(2)

        pyautogui.typewrite("ALL")  # Set Processor to "ALL"
        time.sleep(2)
        pyautogui.press('tab')  # Move to Platform field

        # Set Platform to 'Android'
        pyautogui.press('down')
        time.sleep(2)
        pyautogui.typewrite('Android')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)

        # Proceed through the wizard
        pyautogui.press('space')  # Check Pre-Install
        time.sleep(2)
        pyautogui.press('enter')  # Click Next
        time.sleep(2)

        # Add Files step
        print("Clicking 'Add' button...")
        pyautogui.press('tab')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('down')
        time.sleep(2)
        pyautogui.press('enter')

        print(f"Selecting file: {json_file_path}")
        pyautogui.write(json_file_path)  # Type the file path
        pyautogui.press('enter')
        time.sleep(2)

        print("Setting destination on the device...")
        pyautogui.typewrite("/enterprise/usr/")
        pyautogui.press('enter')
        time.sleep(2)

        # Click Next
        pyautogui.press('tab')  # Move to Project Location field
        time.sleep(2)
        pyautogui.press('tab')  # Skip Project Location field
        time.sleep(2)
        pyautogui.press('enter')

        # Click Finish
        pyautogui.press('enter')

        # Close the success dialog
        pyautogui.press('tab')
        time.sleep(2)
        pyautogui.press('enter')

        # MCSTUDIO MUST BE IN FULL SCREEN MODE!!!!!!!!!!!!
        # Edit Pre-Install script
        print("Editing Pre-Install script...")
        pre_install_coordinates = (-1814, 206)  # Pre-Install
        pyautogui.rightClick(pre_install_coordinates)
        time.sleep(2)

        pre_install_edit_coordinates = (-1769, 243)  # Edit option
        pyautogui.click(pre_install_edit_coordinates)
        time.sleep(2)

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.typewrite("sleep 5\ncopy /sdcard/data.ini /enterprise/usr/\\")
        pyautogui.hotkey('ctrl', 's')
        time.sleep(2)

        # Press F7 to build
        pyautogui.press('f7')
        print(f"Package created successfully: {project_name}")
    except Exception as e:
        print(f"Error while creating package: {e}")

# Loop through all projects
for index, row in df.iterrows():
    raw_project_name = row['NewName'].replace("Migration", "").strip()
    migration_folder_name = f"{raw_project_name}Migration"
    migration_folder = os.path.join("D:\\", migration_folder_name)
    
    # Check if the migration folder exists
    if os.path.exists(migration_folder):
        print(f"Skipping already existing migration: {migration_folder_name}")
        continue
    
    folder_path = row['OriginalName']
    json_file_path = os.path.join(r"E:\Jsons\Home\PDA\DK\foetex", folder_path, "6tap.json")
    print(f"Processing project: {migration_folder_name}")
    create_package(migration_folder_name, json_file_path)



