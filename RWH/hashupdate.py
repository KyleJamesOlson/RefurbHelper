import hashlib
import json
import os

SERVER_PATH = r"\\192.168.5.70\SE Stuff\RWH"
VERSIONS_FILE = os.path.join(SERVER_PATH, "versions.json")
files = ['parser.py', 'template.py', 'utils.py']
versions = {
    "parser.py": "1.0.0",  # Update these manually or increment programmatically
    "template.py": "1.0.0",
    "utils.py": "1.0.0"
}

def calculate_sha256(file_path):
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        while chunk := f.read(8192):
            sha256.update(chunk)
    return sha256.hexdigest()

# Generate versions.json
versions_data = {"modules": {}}
for file in files:
    file_path = os.path.join(SERVER_PATH, file)
    if os.path.exists(file_path):
        hash_value = calculate_sha256(file_path)
        versions_data["modules"][file] = {
            "version": versions[file],
            "sha256": hash_value,
            "path": file
        }
    else:
        print(f"File not found: {file_path}")

with open(VERSIONS_FILE, 'w') as f:
    json.dump(versions_data, f, indent=2)
print(f"Updated {VERSIONS_FILE}")