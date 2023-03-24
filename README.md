# TranslationSourceExcelFilesReportGeneratorGI

This is a simple internal tool that takes a folder with *.XLSX files as an input, and produces a report based on it.
It is used to generate a report about the translation files.

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

