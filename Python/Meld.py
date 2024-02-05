# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 11:04:50 2022

@author: lairde
"""

import os
import csv
import sys
import pypdf
from pypdf import PdfMerger, PdfWriter, PdfReader
from pypdf.annotations import FreeText


# Declare constants

# Folder path: the name of the folder you want to meld
dir_path = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2023-24\Exam\Q12'
# Key file: CSV containing information about the files that were melded
key_file_name = 'keyfile.csv'
# Melded fild name: Name of melded file that you want to create
melded_file_name = 'MeldedPDF.pdf'
# Path to melded file name
melded_file_path = dir_path + '\\' + melded_file_name
# Expected number of files per directory. Usually this is 1 or 2, but some students upload many files.
expectedNumFiles=range (1,10)
# Whether to print the student name at the top of their work, or just their ID
Print_name_on_script = False
# Size of file (in MB) above which to print a warning
warning_file_size = 10
# Number of annotations on a page above which to print a warning
warning_annotation_number = 5


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
    if file_counter not in expectedNumFiles:
        print('Warning: Directory \'' + path + '\' contains unexpected number of files', file_counter)
    return file_counter

if os.path.isfile(melded_file_path):
    print("File " + melded_file_path + "already exists; terminating program.")
    sys.exit()
    
# Create a list of files and top-level directories in dir_path
listOfItems = os.listdir(dir_path)

# Remove everything that isn't a directory
listOfDirectories = list(filter(lambda item: os.path.isdir(dir_path+'\\'+item), listOfItems))

# Remove everything that contains the wrong number of files
listOfDirectories = list(filter(lambda item: 1 <= numItems(dir_path, item) <= 2, listOfDirectories))

with open(dir_path + '\\' + key_file_name, 'w', newline='') as keyfile:
    key = csv.writer(keyfile)
    
    #Create an object to merge the PDF files into
    unscaled_merged_pdf = PdfMerger()
    
    #Iterate through the student directories, merging pdf files together and recording their names and lengths in keyfile.csv
    for item in listOfDirectories:
        for file in os.listdir(dir_path + '\\' + item):
            if file.endswith(".pdf"):    #Merge only pdfs. Ignore other files.
                #Warn if file too big:
                file_stats = os.stat(dir_path + '\\' + item + '\\' + file)
                file_size = file_stats.st_size / (1024 * 1024)
                if (file_size > warning_file_size):
                    print('Warning: The following file is uncommonly large (', round(file_size,1),'MB). You may want to optimise it before marking:\n', item, file)
                readpdf = pypdf.PdfReader(dir_path + '\\' + item + '\\' + file)
                #Warn if too many annotations (This happens when students write their work on an iPad, and makes the melded file unwieldy.)
                for page in readpdf.pages:
                    if "/Annots" in page:
                        if (len(page["/Annots"]) > warning_annotation_number):
                            print('Warning: The following file contains a large number of annotations (', len(page["/Annots"]),'). You may want to print it to an image before marking:\n', item, file)
                #Merge pdfs and record their details in the key file
                totalPages = len(readpdf.pages)
                line = [item, file, totalPages]
                key.writerow(line)
                unscaled_merged_pdf.append(dir_path + '\\' + item + '\\' + file)
    unscaled_merged_pdf.write(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    unscaled_merged_pdf.close()
    
    #Scale the pages in MeldedPDF.pdf to all have A4 width
    reader = PdfReader(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    writer = PdfWriter()
    for x in range(len(reader.pages)):
        # Add a page  from reader, but scale it to A4 (i.e. 595 points):
        page = reader.pages[x]
        page.scale_to(595, 595 * float(page.mediabox.top) / float(page.mediabox.right))
        writer.add_page(page)
        print('Processing page', x+1, 'of', len(reader.pages))
    os.remove(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    
    #Iterate again through the student directories, this time adding labels to the merged PDF near the top of each appropriate page.
    PageCounter = 0
    for item in listOfDirectories:
        for file in os.listdir(dir_path + '\\' + item):
            if file.endswith(".pdf"):
                readpdf = pypdf.PdfReader(dir_path + '\\' + item + '\\' + file)
                page = reader.pages[PageCounter]
                
                #Create the label. Lancaster Moodle seems to use two different directory naming conventions, so comment out the wrong one.
                #StudentName = item.split('_')[0]       #Coursework Moodle
                #StudentID = item.split('_')[1]
                StudentName = item.split(' - ')[1]      #Exam Moodle
                StudentID = item.split(' - ')[0]
                ScriptLabel = StudentName + ' ' + StudentID if Print_name_on_script else StudentID
                
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
            
    #Write the scaled annotated PDF as our final output.
    with open(melded_file_path, "xb") as fp:
        writer.write(fp)
    
    