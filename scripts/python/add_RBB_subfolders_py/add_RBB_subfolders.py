import os
import re
import shutil

# Path to the folder containing the files
source_folder = "/Users/stazh/Documents/Daten/RRB_Apriltranche/RRB_Apriltranche_XML_NER"

print(f"Starting to process files in folder: {source_folder}")

# Get all XML files in the source folder
for filename in os.listdir(source_folder):
    if filename.endswith(".xml"):
        print(f"Processing file: {filename}")
        match = re.match(r"(MM_3_\d+)", filename)
        if match:
            subfolder_name = match.group(1)
            subfolder_path = os.path.join(source_folder, subfolder_name)
            
            # Create subfolder if it doesn't exist
            try:
                os.makedirs(subfolder_path, exist_ok=True)
                print(f"Created or verified existence of subfolder: {subfolder_path}")
            except Exception as e:
                print(f"Error creating subfolder {subfolder_path}: {e}")
                continue
            
            # Copy the file into the subfolder
            src_file = os.path.join(source_folder, filename)
            dest_file = os.path.join(subfolder_path, filename)
            try:
                shutil.copy2(src_file, dest_file)
                print(f"Copied {filename} to {dest_file}")
            except Exception as e:
                print(f"Error copying file {filename} to {dest_file}: {e}")
        else:
            print(f"Filename {filename} does not match the expected pattern.")
    else:
        print(f"Skipping non-XML file: {filename}")

print("Processing complete.")
