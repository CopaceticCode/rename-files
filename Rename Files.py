import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

class RenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Renamer")
        self.folder_path = ""

        # Drop zone for folder
        self.drop_label = ttk.Label(root, text="Drag and drop a folder here or use the button below.", relief="sunken", padding=10)
        self.drop_label.pack(pady=10, padx=10, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        # Browse button
        self.browse_button = ttk.Button(root, text="Browse for Folder", command=self.browse_folder)
        self.browse_button.pack(pady=5)

        # Note about rename.xlsx
        self.note_label = ttk.Label(root, 
            text="Ensure 'rename.xlsx' exists in the selected folder with columns (Old and New).\n\n" + 
                 "Only filenames are used (paths are ignored).\n\n" +
                 "Tip: Get filenames using 'dir -Name' in PowerShell.", 
            wraplength=300)
        self.note_label.pack(pady=10, padx=10)

        # Run button
        self.run_button = ttk.Button(root, text="Rename Files", command=self.run_rename, state="disabled")
        self.run_button.pack(pady=10)

        # Result label
        self.result_label = ttk.Label(root, text="", foreground="green")
        self.result_label.pack(pady=10)

    def on_drop(self, event):
        path = event.data.strip("{}")  # Remove curly braces from Windows paths
        if os.path.isdir(path):
            self.folder_path = path
            self.drop_label.config(text=f"Selected Folder: {self.folder_path}")
            self.run_button.config(state="normal")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            self.drop_label.config(text=f"Selected Folder: {self.folder_path}")
            self.run_button.config(state="normal")

    def run_rename(self):
        if not self.folder_path:
            messagebox.showwarning("No Folder", "Please select a folder first.")
            return

        excel_path = os.path.join(self.folder_path, "rename.xlsx")
        if not os.path.exists(excel_path):
            messagebox.showwarning("Missing File", "The folder must contain a file named 'rename.xlsx'.")
            return

        try:
            # Read the Excel file
            df = pd.read_excel(excel_path, header=None if self.is_headerless(df) else 0)
            df.columns = ["old", "new"]

            # Create a dictionary of old names to new names, using only the filenames
            rename_dict = dict(zip(
                df['old'].apply(os.path.basename),
                df['new'].apply(os.path.basename)
            ))

            # Track successful and failed renames
            successful_renames = []
            failed_renames = []

            for old_name, new_name in rename_dict.items():
                old_path = os.path.join(self.folder_path, old_name)
                new_path = os.path.join(self.folder_path, new_name)

                try:
                    if os.path.exists(old_path):
                        if os.path.exists(new_path):
                            failed_renames.append(f"Cannot rename '{old_name}' to '{new_name}' - file already exists")
                            continue
                        os.rename(old_path, new_path)
                        successful_renames.append(f"Renamed '{old_name}' to '{new_name}'")
                    else:
                        failed_renames.append(f"Source file not found: '{old_name}'")
                except Exception as e:
                    failed_renames.append(f"Error renaming '{old_name}': {str(e)}")

            # Show results
            self.result_label.config(text=f"Done! {len(successful_renames)} files renamed successfully.", foreground="green")
            if failed_renames:
                messagebox.showwarning("Failures", f"{len(failed_renames)} files could not be renamed. Check the console for details.")
                print("\nFailed renames:")
                for failure in failed_renames:
                    print(failure)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def is_headerless(self, df):
        # Check if the first row contains non-string data (assuming headers are strings)
        return not all(isinstance(val, str) for val in df.iloc[0])

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = RenameApp(root)
    root.mainloop()
