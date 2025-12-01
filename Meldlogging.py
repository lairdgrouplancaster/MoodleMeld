from typing import Optional

import tkinter as tk

# ================================================================
# Logging Helper
# ================================================================
def log(message: str, status_widget: Optional[tk.Text] = None) -> None:
    """Log message to status_widget if provided, else print."""
    if status_widget:
        status_widget.insert(tk.END, message + "\n")
        status_widget.see(tk.END)
        status_widget.update_idletasks()
    else:
        print(message)