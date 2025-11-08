import os
import csv
import re
from pathlib import Path
from typing import List, Tuple, Optional

import pypdf
from pypdf import PdfWriter, PdfReader
from pypdf.annotations import FreeText

import tkinter as tk

# -----------------------------
# Constants
# -----------------------------
KEY_FILE_NAME = "key_file.csv"
MELDED_FILE_NAME = "melded_PDF.pdf"
OVERWRITE_ALLOWED = True  # Whether to allow overwrite of melded PDF.
EXPECTED_NUM_FILES = range(1, 10)  # Expected number of files submitted by each student
WARNING_FILE_SIZE_MB = 10  # Warn if submitted file is bigger than this
WARNING_ANNOTATION_COUNT = 5  # Warn if PDF contains more than this number of annotations


# -----------------------------
# Global log helper
# -----------------------------
def log(message: str, status_widget: Optional[tk.Text] = None) -> None:
    """Log message to status_widget if provided, else print to console."""
    if status_widget:
        status_widget.insert(tk.END, message + "\n")
        status_widget.see(tk.END)
        status_widget.update_idletasks()
    else:
        print(message)


# -----------------------------
# Helpers
# -----------------------------
def extract_name_id(s: str) -> Tuple[Optional[str], Optional[int]]:
    """
    Extract (name, id) from the two folder name formats used by Moodle:
      "Name_IDnumber_othertext"
      "IDnumber - Name - othertext"

    Name may include letters, spaces, hyphens or apostrophes.
    IDnumber is an integer.

    Returns (name, id) or (None, None) if no pattern matches.
    """
    # Pattern 1: Name_IDnumber_...
    p1 = re.compile(r"^(?P<name>[A-Za-z][A-Za-z' \-]*)_(?P<id>\d+)")
    # Pattern 2: IDnumber - Name - ...
    p2 = re.compile(
        r"^(?P<id>\d+)\s*-\s*(?P<name>[A-Za-z][A-Za-z' \-]*?)(?=\s*-\s*)"
    )

    for pat in (p1, p2):
        m = pat.match(s)
        if m:
            name = m.group("name").strip()
            id_number = int(m.group("id"))
            return name, id_number

    return None, None


def check_number_of_files(root_path: Path, folder: Path, status_widget: Optional[tk.Text] = None) -> int:
    """
    Count files in a folder and print warnings for unusual structure or file counts.
    Returns the number of files (or -1 if subfolders are present).
    Updates progress live in a Tkinter Text widget (status_widget).
    """
    subfolders = [p for p in folder.iterdir() if p.is_dir()]
    if subfolders:
        log(f"⚠️ Warning: Directory '{folder.relative_to(root_path)}' contains subfolders.", status_widget)
        return -1

    num_files = sum(1 for p in folder.iterdir() if p.is_file())
    if num_files not in EXPECTED_NUM_FILES:
        log(f"⚠️ Warning: Directory '{folder.relative_to(root_path)}' contains {num_files} files.", status_widget)
    return num_files


def scale_pdf_to_width(pdf_path: Path, width: float) -> None:
    """
    Scale all pages of a PDF to a specific width (e.g., 595 for A4).
    """
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    for page in reader.pages:
        aspect_ratio = float(page.mediabox.top) / float(page.mediabox.right)
        page.scale_to(width, width * aspect_ratio)
        writer.add_page(page)

    # Overwrite safely
    tmp_path = pdf_path.with_suffix(".tmp.pdf")
    with open(tmp_path, "wb") as tmp_file:
        writer.write(tmp_file)

    tmp_path.replace(pdf_path)


def merge_pdfs(
    parent_path: Path,
    key_file_path: Path,
    student_folders: List[Path],
    status_widget: Optional[tk.Text] = None,
) -> None:
    """
    Merge PDFs into one and record metadata in a key file.
    Updates progress live in a Tkinter Text widget (status_widget).
    """
    with key_file_path.open("w", newline="") as csv_file:
        key_writer = csv.writer(csv_file)
        merged_pdf = PdfWriter()

        for i, folder in enumerate(student_folders, start=1):
            pdf_paths = sorted(folder.glob("*.pdf"))
            if not pdf_paths:
                log(f"⚠️ No PDF files found in {folder.name}. Skipping.", status_widget)
                continue

            for pdf_path in pdf_paths:
                file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
                if file_size_mb > WARNING_FILE_SIZE_MB:
                    log(f"⚠️ {pdf_path.name} is {file_size_mb:.1f} MB (unusually large).", status_widget)

                reader = PdfReader(str(pdf_path))
                for page in reader.pages:
                    annots = page.get("/Annots", [])
                    if "/Annots" in page and len(page["/Annots"]) > WARNING_ANNOTATION_COUNT:
                        log(f"⚠️ {pdf_path.name} has {len(annots)} annotations (unusually many).", status_widget)

                key_writer.writerow([folder.name, pdf_path.name, len(reader.pages)])
                merged_pdf.append(str(pdf_path))

            log(f"[{i}/{len(student_folders)}] ✅ Melded from {folder.name}", status_widget)

        melded_pdf_path = parent_path / MELDED_FILE_NAME
        merged_pdf.write(str(melded_pdf_path))
        merged_pdf.close()


def annotate_pdf(
    pdf_path: Path,
    student_folders: List[Path],
    show_student_names: bool,
    status_widget: Optional[tk.Text] = None,
) -> None:
    """
    Add annotations (student info) to the top of each corresponding page.
    """
    log(f"✅ Scaling {os.path.basename(pdf_path)} to width...", status_widget)
    
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(str(pdf_path)))

    page_index = 0
    for folder in student_folders:
        for read_pdf_path in folder.glob("*.pdf"):
            reader = PdfReader(str(read_pdf_path))
            page = writer.pages[page_index]

            student_name, student_id = extract_name_id(folder.name)

            if student_name and student_id:
                label_text = f"{student_name} {student_id}" if show_student_names else str(student_id)
                annotation = FreeText(
                    text=label_text,
                    rect=(5, float(page.mediabox.top) - 20, 150, float(page.mediabox.top) - 5),
                    font_size="8pt",
                    font_color="ff0000",
                    border_color="ff0000",
                )
                writer.add_annotation(page_index, annotation)
            else:
                log(f"⚠️ Unknown folder name format of '{folder.name}'. Skipping annotation.", status_widget)
            page_index += len(reader.pages)

    with pdf_path.open("wb") as f:
        writer.write(f)

    log(f"Done.", status_widget)


# -----------------------------
# Main entry point
# -----------------------------
def meld(folder_to_meld: str, show_student_names: bool, status_widget: Optional[tk.Text] = None) -> None:
    """
    Melds a folder of PDFs:
      - merges PDFs
      - scales pages to A4 width
      - adds name/ID annotations
      - writes a key file
    Updates progress live in a Tkinter Text widget (status_widget).
    """
    folder_to_meld = Path(folder_to_meld)
    parent_path = folder_to_meld.parent
    melded_pdf_path = parent_path / MELDED_FILE_NAME
    key_file_path = parent_path / KEY_FILE_NAME

    # Validation
    if not folder_to_meld.is_dir():
        log(f"❌ Folder {folder_to_meld} does not exist.", status_widget)
        return f"❌ Folder {folder_to_meld} does not exist."


    # Collect valid student folders
    student_folders = [
        f for f in folder_to_meld.iterdir()
        if f.is_dir() and check_number_of_files(folder_to_meld, f, status_widget) in EXPECTED_NUM_FILES
    ]

    if not student_folders:
        log("⚠️ No valid student folders found.", status_widget)
        return "⚠️ No valid student folders found."

    # Perform operations
    log(f"✅ Melding from {folder_to_meld.resolve()} ...", status_widget)

    if melded_pdf_path.exists():
        if OVERWRITE_ALLOWED:
            log(f"⚠️ Overwriting '{melded_pdf_path.resolve()}'.", status_widget)
        else:
            log(f"❌ File {melded_pdf_path} already exists.", status_widget)
            return f"❌ File {melded_pdf_path} already exists."

    merge_pdfs(parent_path, key_file_path, student_folders, status_widget)
    scale_pdf_to_width(melded_pdf_path, 595)  # A4 width
    annotate_pdf(melded_pdf_path, student_folders, show_student_names, status_widget)

    log(f"✅ Melded {len(student_folders)} folders successfully.", status_widget)