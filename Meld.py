import os
import csv
import re
from pathlib import Path
from typing import List, Tuple, Optional, Callable

import tkinter as tk
from tkinter import messagebox as mb
from pypdf import PdfWriter, PdfReader
from pypdf.annotations import FreeText
from pypdf.generic import RectangleObject

from Meldlogging import log

# ================================================================
# Constants
# ================================================================
KEY_FILE_NAME = "key_file.csv"
MELDED_FILE_NAME = "melded_PDF.pdf"

EXPECTED_NUM_FILES = range(1, 10)

WARNING_FILE_SIZE_MB = 10
WARNING_ANNOTATION_COUNT = 5


# ================================================================
# Filename / Student Info Parsing
# ================================================================
def extract_name_id(text: str) -> Tuple[Optional[str], Optional[int]]:
    """
    Extract (name, id) from folder names:

        Name_IDnumber_other...
        IDnumber - Name - other...

    Returns (name, id) or (None, None).
    """
    patterns = [
        r"^(?P<name>[A-Za-z][A-Za-z' \-]*)_(?P<id>\d+)",                   # Name_ID
        r"^(?P<id>\d+)\s*-\s*(?P<name>[A-Za-z][A-Za-z' \-]*?)(?=\s*-\s*)"  # ID - Name -
    ]

    for pat in patterns:
        m = re.match(pat, text)
        if m:
            return m.group("name").strip(), int(m.group("id"))

    return None, None


# ================================================================
# Directory Validation
# ================================================================
def check_number_of_files(
    root: Path, folder: Path, status_widget: Optional[tk.Text] = None
) -> int:
    """Ensure no subdirectories exist and file count is plausible."""
    subdirs = [p for p in folder.iterdir() if p.is_dir()]
    if subdirs:
        log(f"⚠️ '{folder.relative_to(root)}' contains subfolders.", status_widget)
        return -1

    files = [p for p in folder.iterdir() if p.is_file()]
    if len(files) not in EXPECTED_NUM_FILES:
        log(
            f"⚠️ '{folder.relative_to(root)}' contains {len(files)} files.",
            status_widget,
        )

    return len(files)


# ================================================================
# PDF Scaling
# ================================================================
def scale_pdf_to_width(
    pdf_path: Path,
    target_width: float,
    status_widget: Optional[tk.Text] = None
) -> None:
    """
    Scale all pages of pdf_path so the visible cropbox width equals target_width.
    """

    pdf_path = Path(pdf_path)
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    total = len(reader.pages)
    log(f"➡️ Scaling {pdf_path.name}: {total} pages", status_widget)

    for idx, page in enumerate(reader.pages, start=1):

        # --- Prefer cropbox ---
        try:
            crop = page.cropbox
            left = float(crop.left)
            right = float(crop.right)
            bottom = float(crop.bottom)
            top = float(crop.top)
            crop_width = right - left
            crop_height = top - bottom
        except Exception:
            # fallback to mediabox
            crop_width = float(page.mediabox.right) - float(page.mediabox.left)
            crop_height = float(page.mediabox.top) - float(page.mediabox.bottom)

        # Handle broken boxes
        if crop_width <= 0 or crop_height <= 0:
            try:
                crop_width = float(page.mediabox.width)
                crop_height = float(page.mediabox.height)
            except Exception:
                writer.add_page(page)
                log(f"⚠️ Page {idx}/{total}: invalid box, skipped.", status_widget)
                continue

        aspect = crop_height / crop_width
        target_height = target_width * aspect

        # Scale the page
        page.scale_to(target_width, target_height)

        # Update cropbox
        page.cropbox = RectangleObject([
            0, 0,
            float(target_width),
            float(target_height)
        ])

        writer.add_page(page)
        log(f"   • Page {idx}/{total} scaled", status_widget)

    # Write output safely
    tmp_path = pdf_path.with_suffix(".tmp.pdf")
    with tmp_path.open("wb") as f:
        writer.write(f)
    tmp_path.replace(pdf_path)

    log("✅ PDF scaling complete", status_widget)


# ================================================================
# PDF Merging
# ================================================================
def merge_pdfs(
    parent: Path,
    key_file_path: Path,
    student_folders: List[Path],
    status_widget: Optional[tk.Text] = None,
) -> None:
    """Merge all student PDFs into one and write key file."""
    merged = PdfWriter()

    with key_file_path.open("w", newline="") as csv_file:
        key_writer = csv.writer(csv_file)

        for i, folder in enumerate(student_folders, start=1):
            pdfs = sorted(folder.glob("*.pdf"))
            if not pdfs:
                log(f"⚠️ No PDFs in {folder.name}. Skipping.", status_widget)
                continue

            for pdf in pdfs:
                size_mb = pdf.stat().st_size / 1024**2
                if size_mb > WARNING_FILE_SIZE_MB:
                    log(f"⚠️ {pdf.name} is {size_mb:.1f} MB.", status_widget)

                reader = PdfReader(str(pdf))
                if "/Annots" in reader.pages[0] and len(reader.pages[0]["/Annots"]) > WARNING_ANNOTATION_COUNT:
                    log(f"⚠️ {pdf.name} has many annotations.", status_widget)

                key_writer.writerow([folder.name, pdf.name, len(reader.pages)])
                merged.append(str(pdf))

            log(f"[{i}/{len(student_folders)}] Melded {folder.name}", status_widget)

    merged.write(str(parent / MELDED_FILE_NAME))


# ================================================================
# PDF Annotation
# ================================================================
def annotate_pdf(
    pdf_path: Path,
    student_folders: List[Path],
    show_names: bool,
    status_widget: Optional[tk.Text] = None,
) -> None:
    """Add each student's name/ID at the top of corresponding pages."""
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(str(pdf_path)))

    page_idx = 0
    for folder in student_folders:
        for pdf_file in folder.glob("*.pdf"):
            num_pages = len(PdfReader(str(pdf_file)).pages)
            page = writer.pages[page_idx]

            name, sid = extract_name_id(folder.name)
            if name and sid:
                label = f"{name} {sid}" if show_names else str(sid)
                annot = FreeText(
                    text=label,
                    rect=(5, float(page.mediabox.top) - 20, 150, float(page.mediabox.top) - 5),
                    font_size="8pt",
                    font_color="ff0000",
                    border_color="ff0000",
                )
                writer.add_annotation(page_idx, annot)
            else:
                log(f"⚠️ Bad folder name: {folder.name}", status_widget)

            page_idx += num_pages

    with pdf_path.open("wb") as f:
        writer.write(f)

    log("Annotation complete.", status_widget)


# ================================================================
# Main Entry
# ================================================================
def meld(folder: str, show_student_names: bool, status_widget: Optional[tk.Text] = None) -> None:
    """Top-level operation: merge, scale, annotate."""
    folder = Path(folder)
    parent = folder.parent
    melded_pdf = parent / MELDED_FILE_NAME
    key_file = parent / KEY_FILE_NAME

    # Validate input
    if not folder.is_dir():
        return log(f"❌ Folder does not exist: {folder}", status_widget)

    if melded_pdf.exists():
        try:
            # Open with no truncation to detect locks
            with open(melded_pdf, "ab"):
                pass
        except Exception:
            return log(f"❌ Cannot overwrite '{melded_pdf}': file is currently in use.", status_widget)
        answer = mb.askyesno(
            "Overwrite existing PDF?",
            f"The output file already exists:\n\n{melded_pdf}\n\nOverwrite it?"
        )
        if not answer:
            return log("❌ Meld cancelled by user (file exists).", status_widget)

    # Gather valid student folders
    students = [
        f for f in folder.iterdir()
        if f.is_dir() and check_number_of_files(folder, f, status_widget) in EXPECTED_NUM_FILES
    ]

    if not students:
        return log("⚠️ No valid student folders found.", status_widget)

    students.sort(key=lambda p: p.name.lower())

    log(f"Starting meld from: {folder.resolve()}", status_widget)

    # Operations
    merge_pdfs(parent, key_file, students, status_widget)
    scale_pdf_to_width(melded_pdf, 595, status_widget)
    annotate_pdf(melded_pdf, students, show_student_names, status_widget)

    log(f"✅ Completed: {len(students)} folders processed.", status_widget)
