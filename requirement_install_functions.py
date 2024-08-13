import string
import subprocess
import sys
import os

def clean_string(s):
    # Create a whitelist of printable ASCII characters
    printable = set(string.printable)
    
    # Remove non-printable characters
    cleaned = ''.join(filter(lambda x: x in printable, s)).strip()
    
    # Remove any leading or trailing spaces
    cleaned = cleaned.strip()
    
    # Return the cleaned string
    return cleaned

def install_required_packages(requirement_file):
    with open(requirement_file, 'r') as file:
        required_packages = [clean_string(line.split('==')[0]) for line in file.readlines() if line.strip()]

    for package in required_packages:
        if package:  # Ensure the package name is not an empty string
            print(f"Checking package: '{package}'")
            if not is_package_installed(package):
                print(f"Installing {package}...")
                try:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                    print(f"{package} installed successfully.")
                except Exception as e:
                    print(f"Error installing package {package}: {e}")
                    raise

def is_package_installed(required_package):
    try:
        subprocess.check_output([sys.executable, '-m', 'pip','show', required_package])
        return True
    except subprocess.CalledProcessError:
        return False
