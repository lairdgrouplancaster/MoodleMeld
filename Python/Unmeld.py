# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 18:11:22 2022

@author: lairde
"""

import os, os.path
import csv
#import pypdf
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import FreeText
import shutil

# Declare constants:

# Folder path: the name of the folder you want to meld
DIR_PATH = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2024-5\Worksheets\Worksheet 4\Marking\Lates'
# Melded fild name: File containing marked work that you want to unmeld
MARKED_FILE_NAME = 'MeldedPDF.pdf'
# Key file: CSV containing information about the files that were melded
KEY_FILE_NAME = 'keyfile.csv'
# Your initials
MY_INITIALS = 'EAL'
# Maximum length of file name of unmelded files. Make this number small if you find yourself creating very long path names.
MAX_FILE_NAME_CHARS = 20



def safe_open_w(path):
    ''' Open "path" for writing, creating any parent directories as needed.
    '''
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return open(path, 'wb')

with open(DIR_PATH + '\\' + KEY_FILE_NAME, 'r', newline='') as keyfile:
    reader = csv.reader(keyfile)
    
    # Iterate over each row in the csv file using reader object
    melded_pages_read = 0
    for row in reader:
        with open(DIR_PATH + '\\' + MARKED_FILE_NAME, 'rb') as infile:
            #Extract the appropriate number of pages from the melded PDF
            reader = PdfReader(infile)
            writer = PdfWriter()
            numPages = int(row[2])
            for page_in_output in range(numPages):
                writer.add_page(reader.pages[melded_pages_read + page_in_output])
            melded_pages_read += numPages
            
            # Annotate the first page bottom left, e.g. with the marker's initials
            annotation = FreeText(
                text=MY_INITIALS,  #Marker's initials
                rect=(20, 20, 60, 40),
                font="Arial",
                bold=True,
                font_size="16pt",
                font_color="ff0000",
                border_color="ff0000",
                #background_color="cdcdcd",
                )
            writer.add_annotation(page_number=0, annotation=annotation)
            
            # Write the extracted pages to the correct place within the 'Unmelded' directory:
            output_directory = row[0]
            output_file = row[1]
            # Truncate length to avoid problems with Windows max path length.
            output_file = output_file[:MAX_FILE_NAME_CHARS] + '.pdf'
            with safe_open_w(DIR_PATH + '\\Unmelded\\' + output_directory +'\\' + output_file) as outfile:
                writer.write(outfile)
                outfile.close

# Zip the unmelded directory, ready for upload to Moodle:
shutil.make_archive(DIR_PATH + '\\To_upload', 'zip', DIR_PATH + '\\Unmelded')
