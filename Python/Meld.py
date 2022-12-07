# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 11:04:50 2022

@author: lairde
"""

import os
import csv
import PyPDF2
from PyPDF2 import PdfFileMerger, PdfWriter, PdfReader

#Declare constants

# Folder path: the name of the folder you want to meld
dir_path = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2022\Worksheets\Worksheet 3\Marking\LairdTest'
#Key file: CSV containing information about the files that were melded
KeyFileName = 'keyfile.csv'
#Melded fild name: Name of melded file that you want to create
MeldedFileName = 'MeldedPDF.pdf'


def numItems(root_path, path):
    """
    Returns the number of objects in a directory and prints a warning if this number is not in the expected range.
    root_path: The directory you want to meld
    path:      The directory within root_path that you are looking inside
    """
    count = 0
    # Iterate through directory
    for item in os.listdir(root_path + '\\' + path):
        # Check if current item is a file; if not return -1 to indicate an error.
        if os.path.isfile(os.path.join(root_path + '\\' + path, item)):
            count += 1
        else:
            print('Warning: Directory \'' + path + '\' contains an unexpected sub-folder')
            return -1
    # Warn if the count is outside the expected range.
    # The expected range is 1..2; one file for the submission, possibly one for the ILSP cover sheet
    if (count < 1) or (count > 2):
        print('Warning: Directory \'' + path + '\' contains unexpected number of files', count)
    return count
    
# Create a list of files and top-level directories in dir_path
listOfItems = os.listdir(dir_path)

# Remove everything that isn't a directory
listOfDirectories = list(filter(lambda item: os.path.isdir(dir_path+'\\'+item), listOfItems))

# Remove everything that contains the wrong number of files
listOfDirectories = list(filter(lambda item: 1 <= numItems(dir_path, item) <= 2, listOfDirectories))

with open(dir_path + '\\' + KeyFileName, 'w', newline='') as keyfile:
    writer = csv.writer(keyfile)
    
    #Create anobject to merge the PDF files into
    merger = PdfFileMerger()
    
    #Iterate through the student directories, merging pdf files together and recording their names and lengths in keyfile.csv. Write the output to MeldedPDF_unscaled.pdf
    for item in listOfDirectories:
        for file in os.listdir(dir_path + '\\' + item):
            readpdf = PyPDF2.PdfFileReader(dir_path + '\\' + item + '\\' + file)
            totalPages = readpdf.numPages
            line = [item, file, totalPages]
            writer.writerow(line)
            merger.append(dir_path + '\\' + item + '\\' + file)
    merger.write(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    merger.close()
    
    #Scale the pages to all have A4 width by creating a new scaled PDF
    reader = PdfReader(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    writer = PdfWriter()
    
    for x in range(reader.numPages):
        # add a page  from reader, but scale it to A4 (i.e. 595 points):
        page = reader.pages[x]
        page.scale_to(595, 595 * float(page.mediabox.top) / float(page.mediabox.right))
        writer.add_page(page)
        print('Processing page ', x+1, ' of ', reader.numPages)
        
    os.remove(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    
    #Write the scaled PDF as our final output.
    with open(dir_path + '\\' + MeldedFileName, "wb") as fp:
        writer.write(fp)
        
    
    