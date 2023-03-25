# Translation Report Generation Tool - Chinese GI

This is a small script with GUI that allows you to generate a report based on source translation files.

It takes *.XLSX files as an incoming data. The format of the spreadsheet is simple: header with double-letter language codes, and the source Chinese column as 'CHS'. It can have multiple sheets within one spreadsheet file.

Hardcoded language codes:
    'CHS', 'CHT', 'DE', 'EN', 'ES', 'FR', 'ID', 'JP', 'KR', 'PT', 'RU', 'TH', 'VI', 'TR', 'IT'

Each sheet is treated as an 'ID', and all the content will be calculated of one sheet will be placed in one row.

Description
A tool to generate a translation progress and statistics report based on a folder with *.xlsx files downloaded from *YOUR CAT*. You can then copy this report elsewhere (for example, into the Google Docs spreadsheet).
It runs on your local machine and does not transfer any data over the network.
By default, it takes all the files from \source\ folder of the same folder where the *.EXE is, and puts a report in the \report\ folder.

It takes all the content from 'CHS' column in every Excel spreadsheet, calculates the amount of Chinese Characters (numbers, English words, code, etc. is excluded from the calculation).
For 'Translated' it only calculates the rows where the selected language ('RU' by default) is not empty.
0% progress is colored with light red
100% progress is colored with light green

How to use
1.Download your source files from *YOUR CAT* in the original format:

2.Pick a folder where there are source files, downloaded from yfsail, by clicking 'Browse'
(or just drop files in the \source\ subfolder)

3.Pick a place to save the report, by clicking the button 'Save report to...'

(or don't do anything, it'll save the report in '\reports\ subfolder)

4.Click 'Process files'
It will generate a pre-formatted Excel spreadsheet.


To compile the file into the Windows executable:

pyinstaller --onefile --noconsole --upx-dir "c:\Soft\upx-4.0.2-win64"  --name genshin_tab_counter_3 main.py

