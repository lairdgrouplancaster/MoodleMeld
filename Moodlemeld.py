import os
import traceback
import zipfile
import tempfile
import shutil
import sys
import subprocess
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mb

from Meld import meld
from Unmeld import unmeld
from Meldlogging import log

# ================================================================
# Main Entry
# ================================================================
def moodlemeld(initials: Optional[str] = None) -> None:
    """Create and run the MoodleMeld dialog window."""

    # Remember last directory
    last_dir_file = os.path.join(os.path.expanduser("~"), ".moodlemeld_lastdir")
    if os.path.exists(last_dir_file):
        try:
            with open(last_dir_file, "r", encoding="utf-8") as f:
                last_dir = f.read().strip()
                if not os.path.isdir(last_dir):
                    last_dir = os.getcwd()
        except:
            last_dir = os.getcwd()
    else:
        last_dir = os.getcwd()
    
    # ================================================================
    # Error reporting
    # ================================================================
    def show_error(title: str, err: Exception):
        tb = traceback.format_exc()
        mb.showerror(title, f"{err}\n\nDetails:\n{tb}")
        log(f"Error: {err}", status_text)

        
    # ================================================================
    # Select a folder to be melded
    # ================================================================
    def select_meld():
        nonlocal last_dir
        
        path = fd.askopenfilename(
            title="Choose folder or ZIP file to meld",
            initialdir=last_dir,
            filetypes=(
                ("Folders or ZIP", "*"),
            )
        )

        # ---- Handle selected file or folder ----
        if not path:
            return

        # update memory
        last_dir = os.path.dirname(path)
        with open(last_dir_file, "w", encoding="utf-8") as f:
            f.write(last_dir)
        
        if os.path.isdir(path): # If it's a directory
            folder = path
        elif path.lower().endswith(".zip"): # If it's a ZIP file — extract next to the ZIP
            try:
                zip_dir = os.path.dirname(path)
                base_name = os.path.splitext(os.path.basename(path))[0]
                temp_dir = os.path.join(zip_dir, f"{base_name}_unzipped")
                # If folder exists (from a previous run), wipe it first
                if os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except PermissionError:
                        log(f"⚠ Unzip folder was previously created and is noein use (probably by OneDrive): {temp_dir}", status_text)
                        log(f"To make meld work, delete this folder manually.", status_text)
                        return
                os.makedirs(temp_dir)
    
                with zipfile.ZipFile(path, "r") as z:
                    z.extractall(temp_dir)
                folder = temp_dir
                log(f"📦 Extracted ZIP to: {temp_dir}", status_text)
            except Exception as e:
                show_error("Failed to unzip file", e)
                log(f"❌ Failed to unzip: {e}", status_text)
                return
        else:
            show_error("Invalid selection", "Please choose a folder or ZIP file.")
            return
            filename
            
        # ---- Run meld() ----
        try:
            meld(
                folder,
                show_student_names.get(),
                status_widget=status_text
            )
            log("Done.", status_text)
        except Exception as e:
            show_error("Melding failed", e)
            log(f"❌ Error: {e}", status_text)
    
    
    # ================================================================
    # Select a file to be unmelded
    # ================================================================
    def select_unmeld():
        nonlocal last_dir
        
        filetypes = (('PDF files', '*.pdf'), ('All files', '*.*'))
        filename = fd.askopenfilename(
            title='Choose file to unmeld',
            initialdir=last_dir,
            filetypes=filetypes
        )
        if not filename:
            return

        # update memory
        last_dir = os.path.dirname(filename)
        with open(last_dir_file, "w", encoding="utf-8") as f:
            f.write(last_dir)
        
        # ---- Run unmeld() ----
        try:
            unmeld(
                filename,
                marker_initials.get(),
                zip_unmelded_folder.get(),
                status_widget=status_text
            )
        except Exception as e:
            show_error("Unmelding failed", e)

    
    # ================================================================
    # Open the working directory in Explorer/Finder
    # ================================================================
    def open_working_dir():
        nonlocal last_dir
        if sys.platform == "win32": # Windows
            os.startfile(last_dir)
        elif sys.platform == "darwin": # macOS
            subprocess.run(['open', last_dir], check=True)
        else: # Linux (using xdg-open, common on most distros)
            subprocess.run(['xdg-open', last_dir], check=True) # This is what Gemini recommends; I have no idea whether it works

    
    # ================================================================
    # Build the GUI
    # ================================================================
    def create_widgets():
        # Meld, Unmeld, and Open Working Directory buttons
        ttk.Button(root, text='Meld...', command=select_meld)\
            .grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(root, text='Unmeld...', command=select_unmeld)\
            .grid(row=1, column=2, sticky='ew', padx=5, pady=5)
        ttk.Button(root, text='Open working directory', command=open_working_dir)\
            .grid(row=1, column=3, sticky='ew', padx=5, pady=5)
    
        # Marker initials
        ttk.Label(root, text="Marker initials:").grid(row=2, column=1, sticky='e', pady=(5, 10))
        ttk.Entry(root, width=5, textvariable=marker_initials)\
            .grid(row=2, column=2, sticky='w', pady=(5, 10))
    
        # Checkbuttons
        ttk.Checkbutton(root, text="Show student names on melded file",
                        variable=show_student_names)\
            .grid(row=3, column=1, columnspan=3, sticky='w', padx=5)
        ttk.Checkbutton(root, text="Zip unmelded folder",
                        variable=zip_unmelded_folder)\
            .grid(row=4, column=1, columnspan=3, sticky='w', padx=5)
    
        # Text box and Scrollbar
        scroll = ttk.Scrollbar(root, orient="vertical", command=status_text.yview)
        status_text.configure(yscrollcommand=scroll.set)
    
        status_text.grid(row=6, column=1, columnspan=3, padx=5, pady=10, sticky='nsew')
        scroll.grid(row=6, column=4, sticky='ns')
    
        root.rowconfigure(6, weight=1)

    root = tk.Tk()
    root.title('MoodleMeld')
    root.iconbitmap('MoodleMeld.ico')
    root.geometry('400x300')
    root.resizable(True, True)
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)
    root.columnconfigure(3, weight=1)
    root.columnconfigure(4, weight=0)

    marker_initials      = tk.StringVar(value=initials)
    show_student_names   = tk.BooleanVar(value=True)
    zip_unmelded_folder  = tk.BooleanVar(value=True)

    status_text = tk.Text(root, height=6, width=30, wrap='word')

    create_widgets()

    root.mainloop()

moodlemeld(initials="MKR")