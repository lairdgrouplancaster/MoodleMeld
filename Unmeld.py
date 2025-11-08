import csv
import shutil
from pathlib import Path
from typing import Optional
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import FreeText
import tkinter as tk

# --------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------
KEY_FILE_NAME = "key_file.csv"
UNMELDED_FOLDER_NAME = "Unmelded"
MAX_FILE_NAME_CHARS = 20  # Limit to avoid overly long paths

# --------------------------------------------------------------------
# Logging helper
# --------------------------------------------------------------------
def log(message: str, status_widget: Optional[tk.Text] = None) -> None:
    """Log message to status_widget if provided, else print."""
    if status_widget:
        status_widget.insert(tk.END, message + "\n")
        status_widget.see(tk.END)
        status_widget.update_idletasks()
    else:
        print(message)


# --------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------
def safe_open_w(path: Path):
    """Open `path` for writing in binary mode, creating parent dirs if needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    return open(path, "wb")


def annotate(pdf_path: Path, initials: str) -> None:
    """Annotate the first page of a PDF with the marker's initials (if provided)."""
    if not initials:
        return

    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    annotation = FreeText(
        text=initials,
        rect=(20, 20, 60, 40),
        font="Helvetica",
        bold=True,
        font_size="16pt",
        font_color="ff0000",
        border_color="ff0000",
    )
    writer.add_annotation(page_number=0, annotation=annotation)

    with open(pdf_path, "wb") as f:
        writer.write(f)


# --------------------------------------------------------------------
# Main unmeld function
# --------------------------------------------------------------------
def unmeld(
    pdf_to_unmeld: str,
    initials: str = "",
    zip_unmelded_folder: bool = False,
    status_widget: Optional[tk.Text] = None
) -> None:
    """
    Unmelds a combined (melded) PDF back into individual student PDFs.
    Updates progress live in a Tkinter Text widget (status_widget).
    """
    pdf_path = Path(pdf_to_unmeld)
    parent_path = pdf_path.parent
    key_file_path = parent_path / KEY_FILE_NAME
    unmelded_folder = parent_path / UNMELDED_FOLDER_NAME

    if not pdf_path.exists():
        log(f"❌ Input PDF not found: {pdf_path}", status_widget)
        return
    if not key_file_path.exists():
        log(f"❌ Key file not found: {key_file_path}", status_widget)
        return

    log(f"Starting to unmeld: {pdf_to_unmeld}\n", status_widget)

    with pdf_path.open("rb") as infile:
        reader = PdfReader(infile)
        pages_read = 0

        with key_file_path.open("r", newline="") as keyfile:
            key_reader = list(csv.reader(keyfile))

        total_files = len(key_reader)
        log(f"Found {total_files} student entries in {KEY_FILE_NAME}.\n", status_widget)

        for i, row in enumerate(key_reader, start=1):
            if len(row) < 3:
                log(f"⚠️ Skipping malformed row {i}: {row}", status_widget)
                continue

            student_folder, student_file, file_pages_str = row
            try:
                file_pages = int(file_pages_str)
            except ValueError:
                log(f"⚠️ Invalid page count at row {i}: {file_pages_str}", status_widget)
                continue

            if file_pages <= 0:
                log(f"⚠️ Skipping row {i} with non-positive page count: {file_pages}", status_widget)
                continue

            # Truncate long filenames safely
            stem = Path(student_file).stem[:MAX_FILE_NAME_CHARS]
            safe_name = stem + ".pdf"
            student_file_path = unmelded_folder / student_folder / safe_name

            writer = PdfWriter()
            try:
                for _ in range(file_pages):
                    writer.add_page(reader.pages[pages_read])
                    pages_read += 1
            except IndexError:
                log(f"❌ Not enough pages in melded PDF for row {i} ({student_folder}/{safe_name})", status_widget)
                break

            with safe_open_w(student_file_path) as outfile:
                writer.write(outfile)

            annotate(student_file_path, initials)
            log(f"[{i}/{total_files}] ✅ Created {student_folder}/{safe_name}", status_widget)

    log(f"\n✅ Unmeld complete! Files saved in: {unmelded_folder.resolve()}", status_widget)
        
    if zip_unmelded_folder:
        shutil.make_archive(str(unmelded_folder), "zip", unmelded_folder)
        log(f"\n📦 Files also zipped to: {unmelded_folder}.zip", status_widget)
    
    log(f"\n✅ Done\n", status_widget)

        
        

