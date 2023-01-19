# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 11:04:50 2022

@author: lairde
"""

import os
import csv
import PyPDF2
from PyPDF2 import PdfFileMerger, PdfWriter, PdfReader
from PyPDF2.generic import AnnotationBuilder


#Declare constants

# Folder path: the name of the folder you want to meld
dir_path = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2022\Worksheets\Worksheet 4\Marking\Lates'
#Key file: CSV containing information about the files that were melded
KeyFileName = 'keyfile.csv'
#Melded fild name: Name of melded file that you want to create
MeldedFileName = 'MeldedPDF.pdf'
#Expected number of files per directory. Usually this is 1 or 2, but some students upload many files.
expectedNumFiles=range (1,100)
#Whether to print the student name at the top of their work, or just their ID
PrintNameOnScript = False


def numItems(root_path, path):
    """
    Returns the number of objects in a directory and prints a warning if this number is not in the expected range.
    root_path: The directory you want to meld
    path:      The directory within root_path that you are looking inside
    """
    FileCounter = 0   #Number of files in directory
    # Iterate through directory
    for item in os.listdir(root_path + '\\' + path):
        # Check if current item is a file; if not return -1 to indicate an error.
        if os.path.isfile(os.path.join(root_path + '\\' + path, item)):
            FileCounter += 1
        else:
            print('Warning: Directory \'' + path + '\' contains an unexpected sub-folder')
            return -1
    # Warn if the count is outside the expected range.
    # Usually the count is 1 or 2; one file for the submission, possibly one for the ILSP cover sheet. A very large number may indicate a problem.
    if FileCounter not in expectedNumFiles:
        print('Warning: Directory \'' + path + '\' contains unexpected number of files', FileCounter)
    return FileCounter
    
# Create a list of files and top-level directories in dir_path
listOfItems = os.listdir(dir_path)

# Remove everything that isn't a directory
listOfDirectories = list(filter(lambda item: os.path.isdir(dir_path+'\\'+item), listOfItems))

# Remove everything that contains an unexpected number of files
listOfDirectories = list(filter(lambda item: numItems(dir_path, item) in expectedNumFiles, listOfDirectories))

with open(dir_path + '\\' + KeyFileName, 'w', newline='') as keyfile:
    writer = csv.writer(keyfile)
    
    #Create an object to merge the PDF files into
    merger = PdfFileMerger()
    
    #Iterate through the student directories, merging pdf files together and recording their names and lengths in keyfile.csv. Write the output to MeldedPDF_unscaled.pdf
    for item in listOfDirectories:
        for file in os.listdir(dir_path + '\\' + item):
            readpdf = PyPDF2.PdfFileReader(dir_path + '\\' + item + '\\' + file)
                        
            #Record the page count and merge readPDF into the melded file:
            readPages = readpdf.numPages
            line = [item, file, readPages]
            writer.writerow(line)
            merger.append(dir_path + '\\' + item + '\\' + file)        
    merger.write(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    merger.close()
    
    #Scale the pages to all have A4 width by creating a new scaled PDF
    reader = PdfReader(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    writer = PdfWriter()
    for x in range(reader.numPages):
        # Add a page  from reader, but scaled to A4 width (i.e. 595 points):
        page = reader.pages[x]
        page.scale_to(595, 595 * float(page.mediabox.top) / float(page.mediabox.right))
        writer.add_page(page)
        print('Processing page ', x+1, ' of ', reader.numPages)
    os.remove(dir_path + '\\' +'MeldedPDF_unscaled.pdf')
    
    #Iterate again through the student directories, this time adding labels to the merged PDF near the top of each appropriate page.
    PageCounter = 0
    for item in listOfDirectories:
        for file in os.listdir(dir_path + '\\' + item):
            readpdf = PyPDF2.PdfFileReader(dir_path + '\\' + item + '\\' + file)
            page = reader.pages[PageCounter]
            
            #Create the label
            StudentName = item.split('_')[0]
            StudentID = item.split('_')[1]
            ScriptLabel = StudentName + ' ' + StudentID if PrintNameOnScript else StudentID
            
            #Add the label as an annotation
            annotation = AnnotationBuilder.free_text(
                ScriptLabel,
                rect=(5,  float(page.mediabox.top)-20,150,  float(page.mediabox.top)-5),
                font_size="8pt",
                font_color="ff0000",
                border_color="ffffff",
                )
            writer.add_annotation(page_number=PageCounter, annotation=annotation)
            PageCounter += readpdf.numPages
            
    
    #Write the scaled annotated PDF as our final output.
    with open(dir_path + '\\' + MeldedFileName, "xb") as fp:
       writer.write(fp)
   
   
        
        
    
    