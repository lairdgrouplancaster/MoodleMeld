# -*- coding: utf-8 -*-
"""
Created on Thu Dec  1 18:11:22 2022

@author: lairde
"""

import os, os.path
import csv
#import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder

#Declare constants

# Folder path: the name of the folder you want to meld
dir_path = r'C:\Users\lairde\OneDrive - Lancaster University\OneDrive Documents\Teaching\Lecturing\PHYS102\2022\Exam\Marking\Submitted'
#Melded fild name: File containing marked work that you want to unmeld
MarkedFileName = 'MeldedPDF.pdf'
#Key file: CSV containing information about the files that were melded
KeyFileName = 'keyfile.csv'
#Your initials
MyInitials = 'EAL'
#Maximum length of file name of unmelded files. Make this number small if you find yourself creating very long path names.
MaxFileNameChars = 20



def safe_open_w(path):
    ''' Open "path" for writing, creating any parent directories as needed.
    '''
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return open(path, 'wb')

 
with open(dir_path + '\\' + KeyFileName, 'r', newline='') as keyfile:
    reader = csv.reader(keyfile)
    
    # Iterate over each row in the csv file
    MeldedPagesRead = 0
    for row in reader:
        with open(dir_path + '\\' + MarkedFileName, 'rb') as infile:
            #Extract the appropriate number of pages from the melded PDF
            reader = PdfReader(infile)
            writer = PdfWriter()
            numPages = int(row[2])
            for PageInOutput in range(numPages):
      #           h="h"
                writer.add_page(reader.pages[MeldedPagesRead + PageInOutput])
#                writer.addPage(reader.pages(MeldedPagesRead + PageInOutput))
            MeldedPagesRead += numPages
            
            # Create the annotation and add it
            annotation = AnnotationBuilder.free_text(
                MyInitials,  #Marker's initials
                rect=(20, 20, 60, 40),
                bold=True,
                font_size="16pt",
                font_color="ff0000",
                border_color="ff0000",
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

