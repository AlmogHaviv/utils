import os
import zipfile
import work_ratio


work_ratio.getting_to_the_right_dir("reports")
# Get the current directory
current_directory = os.getcwd()

# Define the name of the zip file you want to create
zip_filename = 'current_folder.zip'

dit_to_zip = current_directory + '\monthly-report-work-ratio'

# Create a zip file and add all the files and subfolders from the current directory
with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for foldername, subfoldername, filenames in os.walk(dit_to_zip):
        for filename in filenames:
            # Calculate the file's full path
            file_path = os.path.join(foldername, filename)
            print(file_path)
            # Calculate the path to store the file inside the zip file
            zip_path = os.path.relpath(file_path, dit_to_zip)
            # Add the file to the zip file
            zipf.write(file_path, zip_path)

print(f'Successfully created {zip_filename}')
