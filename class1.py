import pyautogui
import pygetwindow as gw
import pandas as pd
import time
import os

# Set a global delay
pyautogui.PAUSE = 1.5  # Reduce delay to 1.5 seconds for faster execution

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

# Function to dynamically wait for UI elements
def wait_for_element(image_path, timeout=10):
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateOnScreen(image_path, confidence=0.8)
        if location:
            return location
        time.sleep(0.5)
    return None

# Function to create a single package
def create_package(project_name, json_file_path):
    try:
        print(f"Creating new package: {project_name}")
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(1.5)

        pyautogui.write(project_name)
        pyautogui.press('tab')  # Move to Project Location field
        pyautogui.press('tab')  # Skip Project Location field
        pyautogui.press('tab')  # Move to Processor field

        pyautogui.typewrite("ALL")
        pyautogui.press('tab')

        pyautogui.press('down')
        pyautogui.typewrite('Android')
        pyautogui.press('enter')
        time.sleep(1.5)

        pyautogui.press('enter')
        time.sleep(1.5)

        pyautogui.press('space')
        pyautogui.press('enter')
        time.sleep(1.5)

        print("Clicking 'Add' button...")
        add_button_coordinates = (-808, 396)  # Coordinates for the Add button
        pyautogui.click(add_button_coordinates)
        time.sleep(1.5)

        # Ensure "Add Files" is selected
        print("Selecting 'Add Files' option...")
        add_files_coordinates = (-773, 392)  # Correct coordinates for "Add Files"
        pyautogui.click(add_files_coordinates)
        time.sleep(1.5)

        # File selection dialog
        print(f"Selecting file: {json_file_path}")
        pyautogui.write(json_file_path)
        pyautogui.press('enter')
        time.sleep(1.5)

        # Set Destination on Device
        print("Setting destination on the device...")
        pyautogui.typewrite("/enterprise/usr/")
        ok_button_coordinates = (-703, 339)  # Coordinates for the OK button
        pyautogui.click(ok_button_coordinates)
        time.sleep(1.5)

        # Click Next
        next_button_coordinates = (-934, 637)
        pyautogui.click(next_button_coordinates)
        time.sleep(1.5)

        # Click Finish
        finish_button_coordinates = (-984, 637)  # Updated coordinates for Finish button
        pyautogui.click(finish_button_coordinates)
        time.sleep(1.5)

        # Close the success dialog
        close_button_coordinates = (-832, 561)
        pyautogui.click(close_button_coordinates)
        time.sleep(1.5)

        # Edit Pre-Install script
        print("Editing Pre-Install script...")
        pre_install_coordinates = (-1814, 209)  # Right-click on Pre-Install
        pyautogui.rightClick(pre_install_coordinates)
        time.sleep(1.5)

        pre_install_edit_coordinates = (-1769, 243)  # Left-click Edit option
        pyautogui.click(pre_install_edit_coordinates)
        time.sleep(1.5)

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.typewrite("sleep 5\ncopy /sdcard/data.ini /enterprise/usr/\\")
        pyautogui.hotkey('ctrl', 's')  # Save changes
        time.sleep(1.5)

        # Press F7 to build
        pyautogui.press('f7')
        print(f"Package created successfully: {project_name}")
    except Exception as e:
        print(f"Error while creating package: {e}")

# Loop through all projects
for index, row in df.iterrows():
    project_name = row['NewName']
    folder_path = row['OriginalName']
    json_file_path = os.path.join(r"E:\Jsons\Home\PDA\DK\foetex", folder_path, "6tap.json")
    print(f"Processing project: {project_name}")
    create_package(project_name, json_file_path)
