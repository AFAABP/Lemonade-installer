Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force

try {
    # Attempt to retrieve Python version
    $pythonVersion = & pip --version
    Write-Host "Python is installed. Version: $pythonVersion"
} catch {
    Write-Host "Python is not installed. Attempting to install Python 3.12..."
    # Check for winget availability
    $wingetAvailable = Get-Command "winget" -ErrorAction SilentlyContinue
    if ($wingetAvailable) {
        & winget install Python.Python.3.12
    } else {
        Write-Error "winget is not available. Please install Python 3.12 manually."
        exit
    }
}

# Refreshes PATH environment variable so we can use pip
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Install required packages through pip
pip install pyqt6 pyinstaller pywin32 requests

# Run imagedata.py to generate image_base64.py
& "C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python312\python.exe" imagedata.py

# Check if image_base64.py was generated successfully
if (Test-Path "image_base64.py") {
    Write-Host "image_base64.py generated successfully."
} else {
    Write-Error "Failed to generate image_base64.py. Please check imagedata.py for errors."
    exit
}

# Run PyInstaller to build the applications
pyinstaller --onefile --windowed --icon=lemonade.ico --add-data "lemonade.ico;." installer.py
pyinstaller --onefile --noconsole --icon=lemonade.ico uninstaller.py
