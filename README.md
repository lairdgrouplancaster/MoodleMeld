# MoodleMeld

MoodleMeld is a Python program to speed up marking PDFs from Moodle.
As a marker, when you download a set of submissions from Moodle you will find each one in its own PDF, grouped into one directory for each student. Opening and closing all these PDFs is a pain in the neck.

This program merges all these PDFs into a single file. After you have marked it, it separates them out into their previous directory structure.

## How to install MoodleMeld
### The easiest way
1. Use [GitHub Desktop](https://www.gitkraken.com/lp/github-integration) to clone this repository.
2. On Windows, run the executable `dist/Moodlemeld.exe` to start the graphical interface
3. Practice melding and unmelding the `Sample data` folder, which has the same format as a Moodle download.

### If you don't want to run an .exe file from the internet
1. Download the files in this repository. (You can use GitHub Desktop, but you don't have to.)
2. In a Python environment in which you have installed `pypdf`, start JupyterLab and open the notebook `Demo.ipynb`.
4. Run the first cell to start the graphical interface. (Instructions further down the notebook tell you how to do other things.)
5. Practice melding and unmelding the `Sample data` folder, which has the same format as a Moodle download.

## How to use MoodleMeld
### To meld your submissions:
1. Download your zipped marking folder from Moodle as usual.
2. Start the graphical interface as above.
3. Click on `Meld...`. Select your marking folder.
4. MoodleMeld will now create a file `melded_PDF.pdf`.
5. To see this file, click on `Open working directory'. Copy the file and mark it.
### To unmeld your submissions:
1. Start the graphical interface as above.
2. Click on `Unmeld...`. Select your marked PDF (which needs to be in the same folder as the `keyfile.csv` created during melding).
3. MoodleMeld will now separate your marked PDF into its original directories, but containing your annotations.
### Options
- If you want to see student names (as well as numbers) on the melded PDF, tick "Show student names on melded file" during melding.
- If you want your initials to appear on each marked PDF, enter them in the "Marker Initials" box during unmelding.
- If you want a zipped copy of the unmelded folder, tick "Zip melded folder" during unmelding.
