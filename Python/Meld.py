# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 11:04:50 2022

@author: lairde
"""

import os
import csv
import sys
import pypdf
from pypdf import PdfWriter, PdfReader
from pypdf.annotations import FreeText


# Declare constants:

# Folder path: the name of the folder you want to meld
DIR_PATH = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2024-5\Worksheets\Worksheet 4\Marking\Lates'
# Key file: CSV containing information about the files that were melded
KEY_FILE_NAME = 'keyfile.csv'
# Melded fild name: Name of melded file that you want to create
MELDED_FILE_NAME = 'MeldedPDF.pdf'
# Path to melded file name
MELDED_FILE_PATH = DIR_PATH + '\\' + MELDED_FILE_NAME
# Expected number of files per directory. Usually this is 1 or 2, but some students upload many files.
EXPECTED_NUM_FILES = range (1,10)
# Whether to print the student name at the top of their work, or just their ID
PRINT_NAME_ON_SCRIPT = True
# Size of file (in MB) above which to print a warning
WARNING_FILE_SIZE = 10
# Number of annotations on a page above which to print a warning
WARNING_ANNOTATION_NUMBER = 5


def numItems(root_path, path):
    """
    Returns the number of objects in a directory and prints a warning if this number is not in the expected range.
    root_path: The directory you want to meld
    path:      The directory within root_path that you are looking inside
    """
    file_counter = 0
    # Iterate through directory
    for item in os.listdir(root_path + '\\' + path):
        # Check if current item is a file; if not return -1 to indicate an error.
        if os.path.isfile(os.path.join(root_path + '\\' + path, item)):
            file_counter += 1
        else:
            print('Warning: Directory \'' + path + '\' contains an unexpected sub-folder')
            return -1
    # Warn if the count is outside the expected range.
    # The expected range is 1..2; one file for the submission, possibly one for the ILSP cover sheet
    if file_counter not in EXPECTED_NUM_FILES:
        print('Warning: Directory \'' + path + '\' contains unexpected number of files', file_counter)
    return file_counter

if os.path.isfile(MELDED_FILE_PATH):
    print("File " + MELDED_FILE_PATH + "already exists; terminating program.")
    sys.exit()
    
# Create a list of files and top-level directories in DIR_PATH
listOfItems = os.listdir(DIR_PATH)

# Remove everything that isn't a directory
listOfDirectories = list(filter(lambda item: os.path.isdir(DIR_PATH+'\\'+item), listOfItems))

# Remove everything that contains the wrong number of files
listOfDirectories = list(filter(lambda item: 1 <= numItems(DIR_PATH, item) <= 2, listOfDirectories))

with open(DIR_PATH + '\\' + KEY_FILE_NAME, 'w', newline='') as keyfile:
    key = csv.writer(keyfile)
    
    # Create an object to merge the PDF files into
    unscaled_merged_pdf = PdfWriter()
    
    # Iterate through the student directories, merging pdf files together and recording their names and lengths in keyfile.csv
    for item in listOfDirectories:
        for file in os.listdir(DIR_PATH + '\\' + item):
            if file.endswith(".pdf"):    # Merge only pdfs. Ignore other files.
                # Warn if file too big:
                file_stats = os.stat(DIR_PATH + '\\' + item + '\\' + file)
                file_size = file_stats.st_size / (1024 * 1024)
                if (file_size > WARNING_FILE_SIZE):
                    print('Warning: The following file is uncommonly large (', round(file_size,1),'MB). You may want to optimise it before marking:\n', item, file)
                readpdf = pypdf.PdfReader(DIR_PATH + '\\' + item + '\\' + file)
                # Warn if too many annotations (This happens when students write their work on an iPad, and makes the melded file unwieldy.)
                for page in readpdf.pages:
                    if "/Annots" in page:
                        if (len(page["/Annots"]) > WARNING_ANNOTATION_NUMBER):
                            print('Warning: The following file contains a large number of annotations (', len(page["/Annots"]),'). You may want to print it to an image before marking:\n', item, file)
                # Merge pdfs and record their details in the key file
                totalPages = len(readpdf.pages)
                line = [item, file, totalPages]
                key.writerow(line)
                unscaled_merged_pdf.append(DIR_PATH + '\\' + item + '\\' + file)
    unscaled_merged_pdf.write(DIR_PATH + '\\' +'MeldedPDF_unscaled.pdf')
    unscaled_merged_pdf.close()
    
    # Scale the pages in MeldedPDF.pdf to all have A4 width
    reader = PdfReader(DIR_PATH + '\\' +'MeldedPDF_unscaled.pdf')
    writer = PdfWriter()
    for x in range(len(reader.pages)):
        # Add a page  from reader, but scale it to A4 (i.e. 595 points):
        page = reader.pages[x]
        page.scale_to(595, 595 * float(page.mediabox.top) / float(page.mediabox.right))
        writer.add_page(page)
        print('Processing page', x+1, 'of', len(reader.pages))
    os.remove(DIR_PATH + '\\' +'MeldedPDF_unscaled.pdf')
    
    # Iterate again through the student directories, this time adding labels to the merged PDF near the top of each appropriate page.
    PageCounter = 0
    for item in listOfDirectories:
        for file in os.listdir(DIR_PATH + '\\' + item):
            if file.endswith(".pdf"):
                readpdf = pypdf.PdfReader(DIR_PATH + '\\' + item + '\\' + file)
                page = reader.pages[PageCounter]
                
                #Create the label. Lancaster Moodle seems to use two different directory naming conventions, so comment out the wrong one.
                StudentName = item.split('_')[0]       #Coursework Moodle
                StudentID = item.split('_')[1]
                #StudentName = item.split(' - ')[1]      #Exam Moodle
                #StudentID = item.split(' - ')[0]
                ScriptLabel = StudentName + ' ' + StudentID if PRINT_NAME_ON_SCRIPT else StudentID
                
                #Add the label as an annotation
                annotation = FreeText(
                    text=ScriptLabel,
                    rect=(5,  float(page.mediabox.top)-20,150,  float(page.mediabox.top)-5),
                    font_size="8pt",
                    font_color="ff0000",
                    border_color="ff0000",
                    )
                writer.add_annotation(page_number=PageCounter, annotation=annotation)
                PageCounter += len(readpdf.pages)
            
    # Write the scaled annotated PDF as our final output.
    with open(MELDED_FILE_PATH, "xb") as fp:
        writer.write(fp)
