# File: main.py
# Author: kk3k02
# Date: March 2024


from tkinter import *
from DriveToExcel import ExportToExcel
import threading


class App:
    def __init__(self):
        self.root = root
        root.title("Google Drive to Excel")
        root.geometry('410x180')

        # Frame - top row
        frame_top = Frame(root)
        frame_top.grid(row=0, column=0)

        # Label - Export Google Drive to Excel
        export_label = Label(frame_top, text="Export Google Drive to Excel")
        export_label.grid(row=0, column=0, pady=5, padx=5)

        # Frame - entry/label link
        entry_frame = Frame(frame_top)
        entry_frame.grid(row=1, column=0)

        # Label - Link
        link_label = Label(entry_frame, text="Link: ")
        link_label.grid(row=0, column=0)

        # Entry - Google Drive address
        self.link_entry = Entry(entry_frame)
        self.link_entry.grid(row=0, column=1)
        self.link_entry.bind('<KeyRelease>', self.check_entry)  # Bind the KeyRelease event to the entry

        # Button - Generate Excel file
        self.generate_button = Button(frame_top, text="Generate", command=self.start, state=DISABLED)
        self.generate_button.grid(row=2, column=0, pady=10)

        # Label - printing info about processed files
        self.label_processed = Label(frame_top, text="Ready to start...", bg="white", width=30)
        self.label_processed.grid(row=0, column=1, padx=(15, 0), pady=10, rowspan=3, sticky='nsew')

        # Label - Author
        author_label = Label(root, text="Author:\nJakub Jakubowicz\nkubajak@outlook.com")
        author_label.grid(row=1, column=0, pady=(5, 0), sticky='ns')

        # Bind the close window function to the window close button
        root.protocol("WM_DELETE_WINDOW", self.close_window)

    # Function to check if the entry is filled
    def check_entry(self, event):
        if self.link_entry.get():
            self.generate_button.config(state=NORMAL)
        else:
            self.generate_button.config(state=DISABLED)

    # Start collecting data and exporting to excel
    def start(self):
        folder_link = self.link_entry.get()
        drive = ExportToExcel(folder_link, self.label_processed)
        drive_thread = threading.Thread(target=drive.generate)
        drive_thread.start()

    # Function to close the window
    def close_window(self):
        self.root.quit()


if __name__ == '__main__':
    root = Tk()
    app = App()
    root.resizable(False, False)
    root.mainloop()
