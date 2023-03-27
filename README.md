# Translation Report Generation Tool - Chinese GI

# What is it
This tool helps the translation manager to prepare the report based on the translation source XLSX files.

The output XLSX report helps visualize:
* Total number of source CJK characters for translation
* How are they distributed between files
* How many of strings are already translated

This tool is currently compatible with 'yfsail'-exported files that include translations.

## Technologies used
It's a Python script with GUI (tkinter), that uses Pandas for calculations and Excel processing. It's compiled into an executable file with pyinstaller and compressed with UPX.

## How it works
It takes a folder with Excel as an input. Then, it loads each file, takes the content of 'CHS' column of each sheet (that's source), takes the other languages' column (based on what user selects in GUI as the target language), and performs calculations:
* Total CJK characters in source (excluding code, variables, English, etc.)
* Sum of characters in rows without translation (that's the 'Untranslated' number)
* Substraction of untranslated from total (that's 'Translated')
* And % of completeness between two.

Then it generates a report as an Excel file and stores it in the users' desired location.

## How the source file should be formatted
#### Headers
Should have 2-letter language code. The hardcoded codes are as follow:

    'CHS', 'CHT', 'DE', 'EN', 'ES', 'FR', 'ID', 'JP', 'KR', 'PT', 'RU', 'TH', 'VI', 'TR', 'IT'

#### Sheets
It is designed to calculate data for each sheet. Hence, the report will have a name of the source file, and a list of sheet names, with calculation for each of them.

So, for example, the source translation file of a software localization project could have 5 sheets: 
1. welcome_screen
2. menuLists
3. AboutWindow
4. errorMessages
5. SettingsWindow.
Each would have identical headers, some content in 'CHS', and translations in other columns (if available).


## How to use it

1. Pick a folder where there are source files, downloaded from yfsail, by clicking 'Browse'
(or just drop files in the \source\ subfolder)
2. Pick a place to save the report, by clicking the button 'Save report to...'
(or don't do anything, it'll save the report in '\reports\ subfolder)
3. Select a target language
4. Click 'Process files'

It will generate a pre-formatted Excel spreadsheet.
0% progress is colored with light red
100% progress is colored with light green


## How to compile from source
To compile the file into the Windows executable:

pyinstaller --onefile --noconsole --upx-dir "c:\Soft\upx-4.0.2-win64"  --name genshin_tab_counter_3 main.py

Replace the '--upx-dir' with the actual path to the UPX executable.

