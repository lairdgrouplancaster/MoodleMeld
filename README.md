# MoodleMeld

MoodleMeld is a pair of Python scripts to speed up marking PDFs from Moodle.
As a marker, when you download a set of submissions from Moodle you will find each one in its own PDF, grouped into one directory for each student. Open and closing all these PDFs is annoying.

- The script 'Meld.py' merges all these PDFs into a single file.

- The script 'Unmeld.py' separates them back out into their previous directory structure.


## How to use MoodleMeld
1. Download the work to be marked from Moodle.
2. In 'Meld.py', edit the dir_path string to point to the directory work is stored.
3. Run 'Meld.py'
4. Mark the combined PDF that gets created. (You may want to change its name, to prevent an accidental overwrite.)
5. If you used Acrobat, it seems to be necessary to print the file to PDF to ensure that your annotations get unpacked properly. The new file is the one you should unmeld.
6. In 'Unmeld.py', make sure that dir_path and MarkedFileName point to the marked file.
7. In 'Unmeld.py', check MyInitials is correct.
8. Run 'Unmeld.py'
9. Compress and upload the unmelded directory to Moodle as normal.
