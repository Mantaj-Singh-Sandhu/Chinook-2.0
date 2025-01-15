import os
import subprocess
import sys
import urllib.request

def download_python_installer(version="3.12.8"):
    """
    Download the Python installer for Windows.
    """
    python_url = f"https://www.python.org/ftp/python/{version}/python-{version}-amd64.exe"
    installer_path = f"python-{version}-amd64.exe"

    print(f"Downloading Python {version} installer...")
    urllib.request.urlretrieve(python_url, installer_path)
    print(f"Downloaded Python installer to {installer_path}")
    return installer_path

def install_python(installer_path):
    """
    Install Python silently using the downloaded installer.
    """
    print(f"Installing Python from {installer_path}...")
    # Install silently and add Python to PATH
    result = subprocess.run(
        [
            installer_path,
            "/quiet",
            "InstallAllUsers=1",
            "PrependPath=1"
        ],
        shell=True
    )
    if result.returncode == 0:
        print("Python installation successful.")
    else:
        print("Python installation failed.")
        sys.exit(1)

def install_library(library_name):
    """
    Install a Python library using pip.
    """
    print(f"Installing library: {library_name}...")
    result = subprocess.run([sys.executable, "-m", "pip", "install", library_name])
    if result.returncode == 0:
        print(f"Library '{library_name}' installed successfully.")
    else:
        print(f"Failed to install library '{library_name}'.")

def main():
    # Specify Python version and libraries to install
    python_version = "3.12.8"
    libraries = ["numpy", "pandas", "PyQt5", "openpyxl", "sqlalchemy", "PyAthena", "pyarrow","fastparquet"]

    # Step 1: Download Python installer
    installer_path = download_python_installer(version=python_version)

    # Step 2: Install Python
    install_python(installer_path)

    # Step 3: Install libraries
    for library in libraries:
        install_library(library)

    # Clean up
    if os.path.exists(installer_path):
        os.remove(installer_path)
        print(f"Removed installer: {installer_path}")

    print("Setup complete!")

if __name__ == "__main__":
    main()
