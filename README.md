# MoodleMeld

MoodleMeld is a pair of Python scripts to speed up marking PDFs from Moodle.
As a marker, when you download a set of submissions from Moodle you will find each one in its own PDF, grouped into one directory for each student. Opening and closing all these PDFs is annoying.

- The script `Meld.py` merges all these PDFs into a single file.

- The script `Unmeld.py` separates them back out into their previous directory structure.


## How to use MoodleMeld
1. Download the submissions to be marked from Moodle.
2. In `Meld.py`:
    1. Edit the `dir_path` string to point to the directory where the submissions are stored.
4. Run `Meld.py`
5. If you get warnings, consider optimising your input files or printing them to PDF as images. Then return to step 4.
6. Mark the combined PDF that gets created.
7. If you marked in Acrobat, it seems to be necessary to print the file to PDF to ensure that your annotations get unpacked properly. The new file is the one you should unmeld.
8. In `Unmeld.py`:
    1. Make sure that `dir_path` and `MarkedFileName` point to the marked file.
    2. Check `MyInitials` is correct.
9. Run `Unmeld.py`
10. Compress and upload the unmelded directory to Moodle as normal.
