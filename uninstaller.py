import os
import shutil
import winreg as reg
import sys
import tempfile


# Paths to the shortcuts
desktop_shortcut = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'Lemonade.lnk')
start_menu_shortcut = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Lemonade.lnk')

# Remove the desktop shortcut if it exists
if os.path.exists(desktop_shortcut):
    os.remove(desktop_shortcut)

# Remove the start menu shortcut if it exists
if os.path.exists(start_menu_shortcut):
    os.remove(start_menu_shortcut)


# Construct the path to the application directory dynamically
application_directory = os.path.join(os.environ['LOCALAPPDATA'], 'Lemonade')

# Function to delete contents of the application directory except the uninstaller
def delete_contents(dir_path, skip_file):
    for item in os.listdir(dir_path):
        item_path = os.path.join(dir_path, item)
        if item_path == skip_file:
            continue  # Skip the uninstaller executable
        if os.path.isfile(item_path) or os.path.islink(item_path):
            os.unlink(item_path)
        elif os.path.isdir(item_path):
            shutil.rmtree(item_path)

# Remove the application directory contents except the uninstaller
uninstaller_path = sys.executable
if os.path.exists(application_directory):
    delete_contents(application_directory, uninstaller_path)


# Remove the registry entry
key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Lemonade"
reg.DeleteKey(reg.HKEY_CURRENT_USER, key_path)

# Function to create a batch script for self-deletion
def create_self_deletion_script():
    # Temporary file for the batch script
    fd, batch_script = tempfile.mkstemp(suffix='.bat', text=True)
    with os.fdopen(fd, 'w') as batch:
        batch.write(f"@echo off\n")
        batch.write(f"timeout /t 5 /nobreak > NUL\n")  # Wait for 5 seconds to ensure the uninstaller has terminated
        batch.write(f"del /f /q \"{sys.executable}\"\n")  # Delete the uninstaller executable
        batch.write(f"rmdir /s /q \"{os.path.dirname(sys.executable)}\"\n")  # Delete the uninstaller's directory
        batch.write(f"del \"%~f0\"\n")  # Delete the batch script itself
    return batch_script

# Schedule the uninstaller and its directory for deletion
batch_script = create_self_deletion_script()
os.system(batch_script)

print("Uninstallation Complete. The system will now clean up.")
