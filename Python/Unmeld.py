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

#Declare constants

# Folder path: the name of the folder you want to meld
dir_path = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2022\Laird'
#Melded fild name: File containing marked work that you want to unmeld
MarkedFileName = 'MeldedPDF.pdf'
#Key file: CSV containing information about the files that were melded
KeyFileName = 'keyfile.csv'
#Your initials
MyInitials = 'EAL'
#Maximum length of file name ofr unmelded files. Make this number small if you find yourself creating very long path names.
MaxFileNameChars = 20



def safe_open_w(path):
    ''' Open "path" for writing, creating any parent directories as needed.
    '''
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return open(path, 'wb')

 
with open(dir_path + '\\' + KeyFileName, 'r', newline='') as keyfile:
    reader = csv.reader(keyfile)
    
    # Iterate over each row in the csv file using reader object
    MeldedPagesRead = 0
    for row in reader:
        with open(dir_path + '\\' + MarkedFileName, 'rb') as infile:
            #Extract the appropriate number of pages from the melded PDF
            reader = PdfReader(infile)
            writer = PdfWriter()
            numPages = int(row[2])
            for PageInOutput in range(numPages):
                writer.add_page(reader.pages[MeldedPagesRead + PageInOutput])
            MeldedPagesRead += numPages
            
            # Annotate the first page bottom left, e.g. with the marker's initials
            annotation = FreeText(
                text=MyInitials,  #Marker's initials
                rect=(20, 20, 60, 40),
                font="Arial",
                bold=True,
                font_size="16pt",
                font_color="ff0000",
                border_color="ff0000",
                #background_color="cdcdcd",
                )
            writer.add_annotation(page_number=0, annotation=annotation)
            
            #Write the extracted pages to the correct place within the 'Unmelded' directory
            OutputDirectory = row[0]
            OutputFile = row[1]
            #Truncate length to avoid problems with Windows max path length.
            OutputFile = OutputFile[:MaxFileNameChars] + '.pdf'
            with safe_open_w(dir_path + '\\Unmelded\\' + OutputDirectory +'\\' + OutputFile) as outfile:
                writer.write(outfile)
                outfile.close

