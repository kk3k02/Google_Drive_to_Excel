# File: DriveToExcel.py
# Author: kk3k02
# Date: March 2024

import pandas as pd
from googleapiclient.errors import HttpError
from Google import Create_Service


class ExportToExcel:
    def __init__(self, folder_link, label):
        self.label = label  # GUI Label for showing info and file processing progress
        self.drive_link = folder_link  # Link to Google Drive dir
        self.dir_id = None  # Google Drive dir ID from link
        self.service = None  # Service for Google Drive API
        self.all_files = None  # Store all files fetched from Google Drive
        self.file_data = []  # Files data
        self.dir_name = None  # Name of the main dir
        self.tree = None  # Google Drive files tree

    def generate(self):
        try:
            # API settings
            API_NAME = "drive"
            API_VERSION = 'v3'
            SCOPES = ['https://www.googleapis.com/auth/drive']
            CLIENT_SECRET = "client_secret.json"

            # Create Google Drive API service
            self.service = Create_Service(CLIENT_SECRET, API_NAME, API_VERSION, SCOPES)

            # Set the info label
            info = "Processing files..."
            self.label.config(text=info)

            parts = self.drive_link.split('/')
            self.dir_id = parts[-1].split('?')[0]

            main_dir = self.service.files().get(fileId=self.dir_id, fields="name").execute()
            self.dir_name = main_dir.get('name')

            # Fetch all files and folders
            self.all_files = self.get_all_files(self.service, self.dir_id)

            # Set the info label
            info = "Generating directory tree..."
            self.label.config(text=info)

            # Generating directory tree
            self.tree = self.get_folder_structure(self.dir_id, tree=None)

            counter = 1
            total_files = len(self.all_files)

            # Process all files from directory
            for file in self.all_files:
                # Process current file
                self.process(file)

                # Count the progress of processed files
                progress = round((counter / total_files * 100), 2)

                # Generate the processed file info
                file_name_info = "File Name: {}\n".format(file.get('name'))
                info = "{}Total files: {}\nProcessed files: {}/{}\nProgress: {}%".format(file_name_info,
                                                                                         total_files,
                                                                                         counter,
                                                                                         total_files, progress)
                # Set the info label
                self.label.config(text=info)

                counter += 1

            # Create data frame and saving it to excel
            df = pd.DataFrame(self.file_data)
            excel_filename = 'Google_Drive_Files.xlsx'
            df.to_excel(excel_filename, index=False)
            print(f"Saved to Excel: {excel_filename}")

            # Set the info label
            self.label.config(text='Done!')

        except HttpError as e:
            print(f"Error has occurred: {e}")
            info = "Error!"

            # Set the info label
            self.label.config(text=info)

    # Process the file
    def process(self, file):
        # Building the full path to the file on drive
        full_path = get_full_path(self.tree, file)
        path = self.dir_name
        if full_path:
            path = self.dir_name + full_path

        # Get the modification date
        modified_time = file.get('modifiedTime')

        # Get the owner info
        owner = None
        if 'owners' in file:
            owner = file['owners'][0]['displayName']

        # Get the file size
        file_size = file.get('size')

        # Collecting all file data to the list
        file_info = {
            'Name': file.get('name'),
            'Link': f"https://drive.google.com/file/d/{file['id']}",
            'Path': path,
            'Type': file.get('mimeType'),
            'Owner': owner,
            'Modification Date': modified_time,
            'File Size': file_size
        }
        self.file_data.append(file_info)

    # Get all files and dirs from selected Google Drive dir
    def get_all_files(self, service, folder_id):
        all_files = []
        query = f"'{folder_id}' in parents and trashed=false"
        page_token = None

        while True:
            response = service.files().list(q=query,
                                            fields="nextPageToken, files(id, name, mimeType, modifiedTime, "
                                                   "owners, size, parents)",
                                            pageToken=page_token).execute()
            files = response.get('files', [])
            all_files.extend(files)

            for file in files:
                if file['mimeType'] == 'application/vnd.google-apps.folder':
                    # If it's a folder, recursively call get_all_files
                    all_files.extend(self.get_all_files(service, file['id']))

            page_token = response.get('nextPageToken')
            if not page_token:
                break

        return all_files

    # Generating directory tree
    def get_folder_structure(self, folder_id, tree=None):
        if tree is None:
            tree = []

        # Fetch all files if not already fetched
        if self.all_files is None:
            self.all_files = self.get_all_files(self.service, folder_id)

        # Build the tree structure from all files and folders
        for file in self.all_files:
            if file['parents'][0] == folder_id:  # Check if the file belongs to the current folder
                file_id = file['id']
                file_name = file['name']
                file_type = file['mimeType']

                if file_type == 'application/vnd.google-apps.folder':
                    folder_structure = {'name': file_name, 'children': []}
                    folder_structure['children'] = self.get_folder_structure(file_id, folder_structure['children'])
                    tree.append(folder_structure)
                else:
                    tree.append({'name': file_name, 'type': 'file', 'id': file_id})

        return tree


# Get full path to file/dir
def get_full_path(tree, file):
    full_path = ""
    found = False

    # Recursive search file/dir in a tree
    def find_file_in_tree(file_name, current_tree, current_path=""):
        nonlocal found, full_path
        for item in current_tree:
            if item.get('name') == file_name and item.get('type') == 'file':
                found = True
                full_path = current_path + "/" + file_name
                return
            elif item.get('name') == file_name and item.get('type') != 'file':
                found = True
                full_path = current_path + "/" + file_name
            else:
                find_file_in_tree(file_name, item.get('children', []), current_path + "/" + item.get('name'))

    find_file_in_tree(file.get('name'), tree)

    return full_path if found else None
