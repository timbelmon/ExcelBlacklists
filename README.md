# Excel Sort Script 

This Python script is designed to filter Excel files based on a set of blacklisted words. It reads in each Excel file, applies the filter to specified columns for each blacklisted word, and then writes the filtered data to a new Excel file.

## Prerequisites

- Python (This script was written in Python 3.11.2 Please ensure that your Python version is compatible.)
- Python libraries: pandas, os, fnmatch, openpyxl

## How to use

1. Place all the Excel files that you want to filter in the 'input' directory.
2. Create a 'blacklists' directory and inside it, place text files (.txt) with the names being the column names you want to filter, and the contents being the blacklisted words (one word per line).
3. Run the script. The script will process each Excel file in the 'input' directory, filter the specified columns, and save the filtered data to a new Excel file in the 'output' directory.

## Functionality

- The `match_with_wildcards` function checks if a word matches with any of the patterns in the blacklist.
- The `auto_adjust_columns` function adjusts the width of each column in the Excel file based on the maximum length of the data in that column.
- The script reads the blacklist files from the 'blacklists' directory and stores the blacklisted words in a dictionary.
- The script reads each Excel file from the 'input' directory, filters the specified columns, and writes the filtered data to a new Excel file in the 'output' directory.
- The max amount of Excel rows tested was about 30k. -> Works but took about 1 minute to filter.

~~Please note that the script does not handle special characters like 'ä', 'ö', 'ü' due to the default encoding in Python. To handle these characters, you should set the encoding to 'utf-8' when reading and writing files. This is already done in the script.~~ -> It can handle 'utf-8' now!
