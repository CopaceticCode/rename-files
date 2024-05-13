import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

# Define color palette
BG_COLOR = "#2c2f33"
FG_COLOR = "#f0f0f0"
ACCENT_COLOR = "#0097e6"

def rename_files():
    # Get the selected folder and Excel file
    folder_path = folder_entry.get()
    excel_file_path = excel_entry.get()

    # Check if folder and file paths are provided
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder.")
        return
    if not excel_file_path:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path)

        # Extract old and new file names
        old_names = df['old'].tolist()
        new_names = df['new'].tolist()

        # Rename files
        for old_name, new_name in zip(old_names, new_names):
            old_file_path = os.path.join(folder_path, old_name)
            new_file_path = os.path.join(folder_path, new_name)
            os.rename(old_file_path, new_file_path)

        messagebox.showinfo("Success", "Files renamed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, folder_path)

def select_excel_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

# Create the main window
root = tk.Tk()
root.title("File Renamer")
root.geometry("400x250")
root.configure(bg=BG_COLOR)

# Create a label and entry for selecting folder
folder_label = tk.Label(root, text="Select Folder:", fg=FG_COLOR, bg=BG_COLOR)
folder_label.grid(row=0, column=0, padx=5, pady=5)
folder_entry = tk.Entry(root, width=30, fg=FG_COLOR, bg=BG_COLOR)
folder_entry.grid(row=0, column=1, padx=5, pady=5)
folder_button = tk.Button(root, text="Browse", command=select_folder, fg=FG_COLOR, bg=ACCENT_COLOR)
folder_button.grid(row=0, column=2, padx=5, pady=5)

# Create a label and entry for selecting Excel file
excel_label = tk.Label(root, text="Select Excel File:", fg=FG_COLOR, bg=BG_COLOR)
excel_label.grid(row=1, column=0, padx=5, pady=5)
excel_entry = tk.Entry(root, width=30, fg=FG_COLOR, bg=BG_COLOR)
excel_entry.grid(row=1, column=1, padx=5, pady=5)
excel_button = tk.Button(root, text="Browse", command=select_excel_file, fg=FG_COLOR, bg=ACCENT_COLOR)
excel_button.grid(row=1, column=2, padx=5, pady=5)

# Create a button to start renaming files
rename_button = tk.Button(root, text="Rename Files", command=rename_files, fg=FG_COLOR, bg=ACCENT_COLOR)
rename_button.grid(row=2, column=0, columnspan=3, padx=5, pady=10)

# Information at the bottom
info_label = tk.Label(root, text="Select a folder that contains files with names to be changed.\nThe selected Excel file should have headers for two columns 'old' and 'new' with names of the existing files in 'old' and new name in 'new'.\nClick 'Rename Files' to perform the operation.", fg=FG_COLOR, bg=BG_COLOR)
info_label.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

# Centering horizontally
for i in range(3):
    root.grid_columnconfigure(i, weight=1)

# Start the Tkinter event loop
root.mainloop()
