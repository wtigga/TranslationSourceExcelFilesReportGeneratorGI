import pandas as pd
from typing import Dict, List
import os
import glob
import numpy as np
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import unicodedata

#### SOURCE FILE LOAD & PREPARATION ####

def list_xlsx_files(folder_location: str) -> list:  #to get a list of xlsx files
    # Combine folder location with the file pattern
    file_pattern = os.path.join(folder_location, "*.xlsx")

    # Use glob to find all matching files
    xlsx_files = glob.glob(file_pattern)

    return xlsx_files

def source_file_name(source_file_path: str) -> str:
    # Get the file name with the extension
    filename_with_extension = os.path.basename(source_file_path)

    # Remove the file extension
    filename_without_extension, _ = os.path.splitext(filename_with_extension)

    return filename_without_extension

def load_excel_file(source_file_path: str):
    try:
        # Read the Excel file using pandas with sheet_name=None to get all sheets
        df = pd.read_excel(source_file_path, engine='openpyxl', sheet_name=None)
        return df
    except FileNotFoundError:
        print(f"Error: File not found at {source_file_path}")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def extract_column_from_sheets(df: Dict[str, pd.DataFrame], source_lang_code: str) -> Dict[str, List]:
    #It puts all the source lang data into the dictionary, one key per sheet.
    sheet_dict = {}

    # Iterate through all the sheets in the Excel file
    for sheet_name, sheet_df in df.items():
        try:
            # Extract the desired column into a list, excluding the header
            column_data = sheet_df[source_lang_code].dropna().tolist()

            # Add the sheet name and column data to the dictionary
            sheet_dict[sheet_name] = column_data
        except KeyError:
            print(f"Error: Column '{source_lang_code}' not found in sheet '{sheet_name}'")
        except Exception as e:
            print(f"Error: {e}")

    return sheet_dict

def filter_columns(df: Dict[str, pd.DataFrame], source_lang_code: str, target_lang_code: str) -> Dict[str, pd.DataFrame]:
    filtered_sheets = {}
    #return dataframe with just chs and ru
    for sheet_name, sheet_df in df.items():
        try:
            # Keep only the desired columns
            filtered_df = sheet_df[[source_lang_code, target_lang_code]]
            filtered_sheets[sheet_name] = filtered_df
        except KeyError:
            print(f"Error: Columns '{source_lang_code}' or '{target_lang_code}' not found in sheet '{sheet_name}'")
        except Exception as e:
            print(f"Error: {e}")

    return filtered_sheets

#### CALCULATIONS ####

def count_chinese_characters(s):
    count = 0
    for c in s:
        if 'CJK UNIFIED IDEOGRAPH' in unicodedata.name(c, ''):
            count += 1
    return count

'''
def calculate_characters_per_sheet(data_dict: Dict[str, List[str]]) -> Dict[str, int]:
    character_count = {}

    for sheet_name, words in data_dict.items():
        total_characters = sum(len(word) for word in words)
        character_count[sheet_name] = total_characters

    return character_count
'''

def calculate_characters_per_sheet(data_dict: Dict[str, List[str]]) -> Dict[str, int]:
    character_count = {}

    for sheet_name, words in data_dict.items():
        total_chinese_characters = sum(count_chinese_characters(word) for word in words)
        character_count[sheet_name] = total_chinese_characters

    return character_count

'''
def calculate_characters_per_sheet_unique(data_dict: Dict[str, List[str]]) -> Dict[str, int]:
    character_count_unique = {}

    for sheet_name, words in data_dict.items():
        # Remove duplicates from the list
        unique_words = list(set(words))

        # Calculate the total number of characters for unique words
        total_characters = sum(len(word) for word in unique_words)
        character_count_unique[sheet_name] = total_characters

    return character_count_unique'''

def calculate_characters_per_sheet_unique(data_dict: Dict[str, List[str]]) -> Dict[str, int]:
    character_count_unique = {}

    for sheet_name, words in data_dict.items():
        # Remove duplicates from the list
        unique_words = list(set(words))

        # Calculate the total number of Chinese characters for unique words
        total_chinese_characters = sum(count_chinese_characters(word) for word in unique_words)
        character_count_unique[sheet_name] = total_chinese_characters

    return character_count_unique

'''
def character_count_untranslated(data: Dict[str, pd.DataFrame], source: str, target: str) -> Dict[str, int]:
    untranslated_count = {}

    for sheet_name, sheet_df in data.items():
        try:
            # Remove rows where the target_lang_code column is not empty
            untranslated_df = sheet_df[sheet_df[target].isna()]

            # Calculate the total number of characters in the source_lang_code column
            total_characters = untranslated_df[source].str.len().sum()
            untranslated_count[sheet_name] = total_characters
        except KeyError:
            print(f"Error: Columns '{source}' or '{target}' not found in sheet '{sheet_name}'")
        except Exception as e:
            print(f"Error: {e}")

    return untranslated_count
'''

def character_count_untranslated(data: Dict[str, pd.DataFrame], source: str, target: str) -> Dict[str, int]:
    untranslated_count = {}

    for sheet_name, sheet_df in data.items():
        try:
            # Remove rows where the target_lang_code column is not empty
            untranslated_df = sheet_df[sheet_df[target].isna()]

            # Calculate the total number of Chinese characters in the source_lang_code column
            total_chinese_characters = untranslated_df[source].apply(count_chinese_characters).sum()
            untranslated_count[sheet_name] = total_chinese_characters
        except KeyError:
            print(f"Error: Columns '{source}' or '{target}' not found in sheet '{sheet_name}'")
        except Exception as e:
            print(f"Error: {e}")

    return untranslated_count

def translated_character_count(character_count_dict: Dict[str, int], untranslated_character_count: Dict[str, int]) -> Dict[str, int]:
    translated_count = {}

    for sheet_name, total_characters in character_count_dict.items():
        untranslated_characters = untranslated_character_count.get(sheet_name, 0)
        translated_characters = total_characters - untranslated_characters
        translated_count[sheet_name] = translated_characters

    return translated_count

def completion_percentage(input_1, input_2):
    result = {}
    for key in input_1.keys():
        completion = (input_2[key] / input_1[key]) #* 100
        #result[key] = f"{completion:.0f}%"
        result[key] = round(completion, 2)

    return result

### OUTPUT DATA PREPARATIONS ###

def combine_dictionaries(dict_list: List[Dict[str, int]]) -> Dict[str, List[int]]:
    combined_dict = {}

    for dictionary in dict_list:
        for key, value in dictionary.items():
            if key in combined_dict:
                combined_dict[key].append(value)
            else:
                combined_dict[key] = [value]

    return combined_dict

def populate_report_df(data: Dict[str, List[int]], filename: str, report_headers: List[str], existing_df: pd.DataFrame = None) -> pd.DataFrame:
    if existing_df is None:
        # Create an empty DataFrame with the specified headers
        report_df = pd.DataFrame(columns=report_headers)
    else:
        report_df = existing_df.copy()

    # List to store row data
    rows = []

    # Populate the DataFrame
    for sheet_name, values in data.items():
        row_data = [filename, sheet_name] + values

        # Check if the length of the row_data matches the number of headers
        if len(row_data) != len(report_headers):
            row_data += [None] * (len(report_headers) - len(row_data))

        rows.append(row_data)

    # Convert the rows to a DataFrame
    appended_df = pd.DataFrame(rows, columns=report_headers)

    # Append new data to the existing DataFrame
    report_df = pd.concat([report_df, appended_df], ignore_index=True)

    return report_df

def process_excel_files(file_list: list, source_lang_code: str, target_lang_code: str, report_headers: list) -> pd.DataFrame:
    dataframes = []

    for file_path in file_list:
        df = from_file_to_dataframe(file_path, source_lang_code, target_lang_code, report_headers)
        dataframes.append(df)

    # Combine all DataFrames into one
    combined_df = pd.concat(dataframes, ignore_index=True)

    return combined_df

#### EXECUTION OF CALCULATIONS ###

def from_file_to_dataframe(source_file_path, source_lang_code, target_lang_code, report_headers):

    # Load excel file
    original_frame = load_excel_file(source_file_path)

    # Filter columns
    dataframe = filter_columns(original_frame, source_lang_code, target_lang_code)

    # Dict with source only
    data_dict = extract_column_from_sheets(dataframe, source_lang_code)

    # Source characters dict
    character_count_dict = calculate_characters_per_sheet(data_dict)

    # Source unique characters dict
    character_count_dict_unique = calculate_characters_per_sheet_unique(data_dict)

    untranslated_character_count = character_count_untranslated(dataframe, source_lang_code, target_lang_code)

    # Translated characters dict
    translated_count = translated_character_count(character_count_dict, untranslated_character_count)

    # Completion percentage
    completeness = completion_percentage(character_count_dict, translated_count)

    # Create a list of all dicts (except unique)
    full_result_list = [character_count_dict, translated_count, untranslated_character_count, completeness]

    # Create a single dictionary
    full_result_dict = combine_dictionaries(full_result_list)

    # Get the filename without the file extension
    filename = source_file_name(source_file_path)

    # Populate report DataFrame
    report_df = populate_report_df(full_result_dict, filename, report_headers)

    return report_df

### COMPILING AN EXCEL ####

def save_dataframe_to_excel(df: pd.DataFrame, report_save_path: str):
#    df = df.fillna('NA')
#    df = df.replace([np.inf, -np.inf], 'Inf')
    with pd.ExcelWriter(report_save_path, engine='xlsxwriter', options={'nan_inf_to_errors': True}) as writer:
        df.to_excel(writer, index=False, sheet_name="Report", na_rep='')

        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets["Report"]

        # Set column widths
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:Z', 15)

        # Set row height
        for row_num in range(len(df) + 1):
            worksheet.set_row(row_num, 20)  # 30 points height, which is double the standard height

        # Apply vertical and horizontal center alignment, wrap text
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'text_wrap': True
        })
        #################
        # Add the "Total" row
        total_row = len(df) + 1

        # Write the "Total" label
        worksheet.write(total_row, 0, 'Total', cell_format)

        # Write the sum formulas for the 3rd, 4th, and 5th columns
        for col_num in range(2, 5):
            column_letter = chr(ord('A') + col_num)
            worksheet.write_formula(total_row, col_num, f'=SUM({column_letter}2:{column_letter}{total_row})',
                                    cell_format)
        ##################
        # Apply percentage formatting for 6th column
        percentage_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'text_wrap': True,
            'num_format': '0%'
        })

        # Conditional formatting for the 6th column
        red_format = workbook.add_format({'bg_color': '#FFC7CE'})
        green_format = workbook.add_format({'bg_color': '#C6EFCE'})
        worksheet.conditional_format('F2:F{}'.format(len(df) + 1),
                                     {'type': 'cell', 'criteria': '==', 'value': 0, 'format': red_format})
        worksheet.conditional_format('F2:F{}'.format(len(df) + 1),
                                     {'type': 'cell', 'criteria': '==', 'value': 1, 'format': green_format})

        # Write data with formatting
        for row_num, row_data in enumerate(df.values, start=1):
            for col_num, cell_value in enumerate(row_data):
                if col_num == 5:  # Apply percentage formatting to the 6th column (index 5)
                    worksheet.write(row_num, col_num, cell_value, percentage_format)
                else:
                    worksheet.write(row_num, col_num, cell_value, cell_format)

        # Freeze the header row
        worksheet.freeze_panes(1, 0)

        # Set header format
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Calibri',
            'text_wrap': True
        })

        # Write headers with formatting
        for col_num, header in enumerate(df.columns):
            worksheet.write(0, col_num, header, header_format)

        # Merge vertically adjacent cells with the same value only in the first column
        start_row = 1
        for row_num in range(2, len(df) + 2):
            if row_num == len(df) + 1 or df.iat[row_num - 1, 0] != df.iat[row_num - 2, 0]:
                if row_num - 1 > start_row:
                    worksheet.merge_range(start_row, 0, row_num - 1, 0, df.iat[start_row - 1, 0], cell_format)
                start_row = row_num

#save the report in Excel format

def process_and_save(xlsx_files, source_lang_code, target_lang_code, report_headers, report_save_path):
    combined_df = process_excel_files(xlsx_files, source_lang_code, target_lang_code, report_headers)
    save_dataframe_to_excel(combined_df, report_save_path)

#### VARIABLES #####
folder_location = r'C:\1'
folder_location = os.getcwd() + '\source'
xlsx_files = list_xlsx_files(folder_location)

source_file_path = r'c:\1\1Diff-3.7-batch1-beta1-Textmap0317vs0317.xlsx'
source_lang_code = 'CHS'
target_lang_code = 'RU'
report_headers = [
    "file",
    "Key",
    "Chinese Wordcount",
    "Translated",
    "Not_translated",
    "Completeness",
    "Translator",
    "Proofreader",
    "Batch 1",
    "Batch 2",
    "Batch 3",
    "Batch 4",
    "Batch 5",
    "Batch 6",
    "Live"
]
language_codes = [
    'CHS', 'CHT', 'DE', 'EN', 'ES', 'FR', 'ID', 'JP', 'KR', 'PT', 'RU', 'TH', 'VI', 'TR', 'IT'
]
source_lang_codes_all = language_codes
report_save_path = r'c:\2\report.xlsx'
report_save_path = os.getcwd() + (str(r'\reports\report.xlsx'))

#button should call this function

def for_button():
    try:
        process_and_save(xlsx_files, source_lang_code.get(), target_lang_code.get(), report_headers, report_save_path)
        # Show popup window with message "Process complete"
        messagebox.showinfo("Process complete", "Report has been generated. You can find it at: " + str(report_save_path))
    except Exception as e:
        # Show popup window with error message
        messagebox.showerror("Error", str(e))
    print("Button clicked")

#GUI##########################################

# Create a new window object
window = tk.Tk()

# Set the window title
window.title("Genshin Language Source File Statistics")

# Set the window size
window.geometry("1000x300")

# Define a function to browse for a folder
def browse_folder():
    global folder_location
    folder_location = filedialog.askdirectory()
    folder_path_var.set(folder_location)
    global xlsx_files
    xlsx_files = list_xlsx_files(folder_location)

def save_report():
    global report_save_path
    appendix = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=f"report_{appendix}")
    report_save_path_label.config(text=report_save_path)

# Create a frame to hold the browse button and file path
frame = tk.Frame(window)

# Elements for saving report
save_report_button = tk.Button(window, text="Save report to...", command=save_report)
save_report_button.grid(row=0, column=0, padx=10, pady=10)
report_save_path_label = tk.Label(window, text=report_save_path)
report_save_path_label.grid(row=0, column=1, padx=10, pady=10)

# Elements for language codes
lang_codes_label1 = tk.Label(window, text="Source Language Code:")
lang_codes_label1.grid(row=1, column=0, sticky='w', padx=10, pady=10)
source_lang_code = tk.StringVar()
source_lang_combobox = ttk.Combobox(window, textvariable=source_lang_code, values=source_lang_codes_all)
source_lang_combobox.current(source_lang_codes_all.index('CHS'))
source_lang_combobox.grid(row=1, column=1, sticky='w', padx=10, pady=10)

lang_codes_label2 = tk.Label(window, text="Target Language Code:")
lang_codes_label2.grid(row=2, column=0, sticky='w', padx=10, pady=10)
target_lang_code = tk.StringVar()
target_lang_combobox = ttk.Combobox(window, textvariable=target_lang_code, values=source_lang_codes_all)
target_lang_combobox.current(source_lang_codes_all.index('RU'))
target_lang_combobox.grid(row=2, column=1, sticky='w', padx=10, pady=10)

# Create a frame to hold the browse button and file path
frame = tk.Frame(window)
frame.grid(row=3, column=0, padx=10, pady=10)

# Create a button to browse for a folder
browse_button = tk.Button(frame, text="Browse", command=browse_folder)
browse_button.pack(side="left")

# Create a text field to display the file path
folder_path_var = tk.StringVar()
folder_path_var.set(folder_location)
folder_path_entry = tk.Entry(frame, textvariable=folder_path_var, width=80)
folder_path_entry.pack(side="left")

# Button to process files
process_button = tk.Button(window, text="Process Files", command=for_button)
process_button.grid(row=4, column=1, padx=10, pady=10)

# Start the main event loop
window.mainloop()