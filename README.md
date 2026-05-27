# Word Metadata Extractor
This simple script extracts metadata from Word DOCX files in a folder and provides the output as an Excel sheet.

## Usage

### Preparation
Before using this script, ensure all your Word DOCX files are stored in a single folder.

### Running the script
When you run the Python script, it will pop up a file explorer/finder window asking you to select the folder in which your Word DOCX files are stored.

Once you have done that, it will extract the metadata from all Word DOCX files in the chosen folder and save the output to an Excel file in the same folder.

### Output
The script will save the output as an Excel file in the same folder as your Word DOCX files.

Each row in the table represents a Word DOCX file, and the columns are as following:
1. Filename
2. File created (date-time)
3. Title
4. Authors
5. Last saved by
6. Revision Number
7. Content created (date-time)
8. Last date saved (date-time)
9. Total editing time
10. Pages
11. Word count
12. Character count
13. Paragraph count
14. Template
15. (Hidden) Editing time in minutes

There is some formatting within the table:
1. If the Author and Last saver are not the same, then the background is light orange.
2. If the number of revisions is less than 2, then the background is light orange.
3. Total editing time affects the formatting as well, with thirds (33-percentile) formatted differently.
4. If the template is not Normal, then the background is light orange.

### Note
The script uses some standard Python libraries. If you don't have them installed on your system, then in the first run, the script will try to install these dependencies. 

## Caveat
Tested on Windows 11 Education 64-bit

Not tested on Apple iOS or Linux.

## Never run a Python script before?
It's straightforward, but you may need to install Python on your machine first.

### Install Python
Anaconda is one of the most popular distributions of Python. Download and install from https://www.anaconda.com/download

Installation is straightforward, but if you need help, check out https://www.anaconda.com/docs/getting-started/anaconda/install/overview

### Start Spyder
Anaconda comes with Spyder IDE. Start Spyder.

Once Spyder is ready, open the file 'word-metadata-extractor.py' that has the script.

All that's left is for you to hit 'Run', i.e. the green 'Play' button.