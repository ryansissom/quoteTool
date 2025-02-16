import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# List of required packages
required_packages = [
    "xlwings",
    "pandas",
    "fuzzywuzzy"
]

# Install each package
for package in required_packages:
    install(package)

print("All required packages have been installed.")