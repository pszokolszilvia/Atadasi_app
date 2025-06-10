import zipfile
import os

def zip_directory(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, folder_path)
                zipf.write(full_path, arcname=rel_path)
