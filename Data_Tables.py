
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import shutil
import os

# Columns to collect from in Data
section_1     = {"A", "B", "C", "D", "E"}
section_2     = {"A", "B", "C", "D", "E"}
section_3_1   = {"C", "D", "E", "G", "J", "M", "P", "S", "V", "Y", "AB", "AC", "AD", "AE", "AF"}
section_3_1_prev = {"C", "D", "E", "F", "H", "I", "J"}
section_3_1_other = {"C", "D", "E", "G", "J", "M" "P", "S", "V", "Y", "AB", "AC", "AD", "AE", "AF"}
section_3_1_other_prev = {"C", "D", "E", "F", "H", "I", "J"}
section_3_2_A = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W"}
section_3_2_B = {"A", "B", "C", "D", "E"}
section_3_3   = {"C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"}
section_4     = {"A", "B", "C", "D"}
section_5     = {"A", "B", "C", "E", "F", "G", "H", "K", "L", "M", "O", "Q", "R", "S", "T", "U"}
section_6_2   = {"A", "C", "E", "F", "G", "H", "I"}
section_6_3   = 
attach_A      = 
attach_D      = 
attach_E      = 
attach_H      = 


def column_match(label_row: int, data):
    """Selects the "tifnum" column in the excel file

    Args:
        label_row (int): row that holds the labels of the columns (in data)
        data: sheet that contains the input data

    Returns:
        col: "tifnum" column
    """
    for cell in data[label_row]:
        if cell.value:
            label = str(cell.value).strip().lower()
            if label == "tifnum":
                col = cell.column
    return col

def set_data_length(data, label_row: int, row_start: int):
    """Determines the how far down the column to go for the data

    Args:
        data (path): sheet that contains the input data
        label_row (int): row that holds the labels of the columns (in data)
        row_start (int): first row that holds the data we're looking for in the sheet

    Returns:
        length (int): vertical length of data in the sheet
    """
    length = 0
    
    tif_col = column_match(label_row, data)
    row = row_start
    while True:
        cell = data.cell(row=row, column=tif_col)
        if cell:
            length += 1
            row    += 1
        else:
            break
        
    return length

def fill_date(destination, row: int, col: int, reporting_year: str, length: int):
    while length > 0:
        destination.cell(row, col, value=reporting_year)
        row    += 1
        length -= 1
    
    return

def get_column_data(data, col: str, label_row: int, start_row: int, length: int):
    """Collect the data from the input sheet

    Args:
        data (path): sheet that contains the input data
        col (str): lettered column  in sheet where data is being collected like "A" or "B" (must be single column)
        label_row (int): row that holds the labels of the columns (in data)
        start_row (int): first row that holds the data we're looking for in the sheet
        length (int): vertical length of data in the sheet

    Returns:
        values (list): list of values in a column of data
    """
    length = set_data_length(data, label_row, start_row)
    cells  = data[col][start_row-1 : start_row+length]
    values = [cell.value for cell in cells]
    
    return values

# column_labels = mapped list that contains the lowercase names of important columns
# column_map    = mapped list that contains the location of the columns in the output
def populate(data, destination, column_labels, column_map):
    """Populates an entire column of data from one workbook to another

    Args:
        data (path): sheet that contains the input data
        destination (path): sheet that contains the output data
        column_labels (_type_): _description_
        column_map (_type_): _description_
    """
    

def Data_Tables(reporting_year, input_file, template_file):
    shutil.copy('2023 ATR Data Tables - Final 2025 01 03.xlsx', 'Data Tables Copy.xlsx')  # CHANGE IF FILE CHANGES
    copied_table = load_workbook('Data Tables Copy.xlsx')
