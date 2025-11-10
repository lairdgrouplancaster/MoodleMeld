import os
import traceback
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
        folder = fd.askdirectory(
            title='Choose folder to meld',
            mustexist=True,
            initialdir=os.getcwd()
        )
        if not folder:
            return
        try:
            meld(
                folder,
                show_student_names.get(),
                status_widget=status_text
            )
            log("Done.")
        except Exception as e:
            show_error("Melding failed", e)

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

        # Checkbuttons for options
        ttk.Checkbutton(root, text="Show student names on melded file", variable=show_student_names).grid(
            row=3, column=1, columnspan=2, sticky='w', padx=5
        )
        ttk.Checkbutton(root, text="Zip unmelded folder", variable=zip_unmelded_folder).grid(
            row=4, column=1, columnspan=2, sticky='w', padx=5
        )

        # Status text box
        status_text.grid(row=6, column=1, columnspan=2, padx=5, pady=10, sticky='nsew')
        root.rowconfigure(6, weight=1)

    root = tk.Tk()
    root.title('MoodleMeld')
    # TODO: root.iconbitmap('path_to_icon.ico')  # Optional: set window icon
    root.geometry('610x300')
    root.resizable(True, True)
    root.columnconfigure(1, weight=1)
    root.columnconfigure(2, weight=1)

    marker_initials = tk.StringVar()
    show_student_names = tk.BooleanVar()
    zip_unmelded_folder = tk.BooleanVar()

    status_text = tk.Text(root, height=6, width=50, wrap='word')

    show_student_names.set(True)
    zip_unmelded_folder.set(True)
    marker_initials.set(initials)

    create_widgets()

    root.mainloop()
