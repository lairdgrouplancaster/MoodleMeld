# MoodleMeld

MoodleMeld is a Python program to speed up marking PDFs from Moodle.
As a marker, when you download a set of submissions from Moodle you will find each one in its own PDF, grouped into one directory for each student. Opening and closing all these PDFs is a pain in the neck.

This program merges all these PDFs into a single file. After you have marked it, it separates them out into their previous directory structure.

## How to install MoodleMeld
1. Download the files in this repository.
2. In a Python environment in which you have installed `pypdf', start JupyterLab and open the notebook `Demo.ipynb`.
3. Run the first cell to start the graphical interface.
4. Practice melding and unmelding the `Sample data' folder, which has the same format as a Moodle download.

## How to use MoodleMeld
### To meld your submissions:
1. Download your marking folder from Moodle as usual.
2. Start the graphical interface as above.
3. Click on `Meld`. Select your marking folder.
4. MoodleMeld will now create a file `melded_PDF.pdf`. Copy this and mark it.
### To unmeld your submissions:
1. Start the graphical interface as above.
2. Click on `Unmeld`. Select your marked PDF (which needs to be in the same folder as the `keyfile.csv` created during melding).
3. MoodleMeld will now separate your marked PDF into its original directories, but containing your annotations.
### Options
- If you want to see student names (as well as numbers) on the melded PDF, tick "Show student names on melded file" during melding.
- If you want your initials to appear on each marked PDF, enter them in the "Marker Initials" box during unmelding.
- If you want a zipped copy of the unmelded folder, tick "Zip melded folder" during unmelding.
