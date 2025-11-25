import os
import traceback
import zipfile
import tempfile
import shutil
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mb

from Meld import meld
from Unmeld import unmeld

def moodlemeld(initials: Optional[str] = None) -> None:
    """Create and run the MoodleMeld dialog window."""

    def log(message: str, newline=True):
        end_char = "\n" if newline else ""
        status_text.insert(tk.END, message + end_char)
        status_text.see(tk.END)
        status_text.update_idletasks()

    def show_error(title: str, err: Exception):
        tb = traceback.format_exc()
        mb.showerror(title, f"{err}\n\nDetails:\n{tb}")
        log(f"Error: {err}")

    def select_folder():
        # Allow user to select either a folder or a ZIP file
        path = fd.askopenfilename(
            title="Choose folder or ZIP file to meld",
            initialdir=os.getcwd(),
            filetypes=(
                ("Folders or ZIP", "*"),
            )
        )

        # ---- Handle selected file or folder ----
        if not path:
            return
        if os.path.isdir(path): # If it's a directory
            folder = path
        elif path.lower().endswith(".zip"): # If it's a ZIP file — extract next to the ZIP
            try:
                zip_dir = os.path.dirname(path)
                base_name = os.path.splitext(os.path.basename(path))[0]
                temp_dir = os.path.join(zip_dir, f"{base_name}_unzipped")
                # If folder exists (from a previous run), wipe it first
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
    
                with zipfile.ZipFile(path, "r") as z:
                    z.extractall(temp_dir)
                folder = temp_dir
                log(f"📦 Extracted ZIP to: {temp_dir}")
            except Exception as e:
                show_error("Failed to unzip file", e)
                log(f"❌ Failed to unzip: {e}")
                return
        else:
            show_error("Invalid selection", "Please choose a folder or ZIP file.")
            return
            
        # ---- Run meld() ----
        try:
            meld(
                folder,
                show_student_names.get(),
                status_widget=status_text
            )
            log("Done.")
        except Exception as e:
            show_error("Melding failed", e)
            log(f"❌ Error: {e}")

    def select_file():
        filetypes = (('PDF files', '*.pdf'), ('All files', '*.*'))
        filename = fd.askopenfilename(
            title='Choose file to unmeld',
            initialdir=os.getcwd(),
            filetypes=filetypes
        )
        if not filename:
            return
        try:
            unmeld(
                filename,
                marker_initials.get(),
                zip_unmelded_folder.get(),
                status_widget=status_text
            )
        except Exception as e:
            show_error("Unmelding failed", e)

    def create_widgets():
        # Meld and Unmeld buttons
        ttk.Button(root, text='Meld...', command=select_folder).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(root, text='Unmeld...', command=select_file).grid(row=1, column=2, sticky='ew', padx=5, pady=5)
    
        # Marker initials input
        ttk.Label(root, text="Marker initials:").grid(row=2, column=1, sticky='e', pady=(5, 10))
        ttk.Entry(root, width=5, textvariable=marker_initials).grid(row=2, column=2, sticky='w', pady=(5, 10))
    
        # Checkbuttons
        ttk.Checkbutton(root, text="Show student names on melded file",
                        variable=show_student_names).grid(row=3, column=1, columnspan=2, sticky='w', padx=5)
        ttk.Checkbutton(root, text="Zip unmelded folder",
                        variable=zip_unmelded_folder).grid(row=4, column=1, columnspan=2, sticky='w', padx=5)
    
        # --- Status text box + Scrollbar ---
        scroll = ttk.Scrollbar(root, orient="vertical", command=status_text.yview)
        status_text.configure(yscrollcommand=scroll.set)
    
        status_text.grid(row=6, column=1, columnspan=2, padx=5, pady=10, sticky='nsew')
        scroll.grid(row=6, column=3, sticky='ns')   # <-- Add scrollbar in column 3
    
        # Make row 6 resize with window
        root.rowconfigure(6, weight=1)


    root = tk.Tk()
    root.title('MoodleMeld')
    # TODO: root.iconbitmap('path_to_icon.ico')  # Optional: set window icon
    root.geometry('610x300')
    root.resizable(True, True)
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)

    marker_initials      = tk.StringVar(value=initials)
    show_student_names   = tk.BooleanVar(value=True)
    zip_unmelded_folder  = tk.BooleanVar(value=True)

    status_text = tk.Text(root, height=6, width=50, wrap='word')

    create_widgets()

    root.mainloop()
