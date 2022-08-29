#Interface
from tkinter import *
import tkinter.font as font
from tkinter.filedialog import askopenfilename

#Excel
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def file_is_not_valid(dictionary, key):
    if key not in dictionary:
        return True
    return False
    ...


def get_excel_filename(dictionary, key, title='Select a File'):
    while file_is_not_valid(dictionary, key):
        dictionary[key] = askopenfilename(initialdir='.',
                                          title=title,
                                          filetypes=[("Excel files", ".xlsx")])
    return dictionary


def execute_program(files):
    # Variables for source file, worksheets, and empty dictionary for dataframes
    spreadsheet_file = pd.ExcelFile(files['input_filename'])
    worksheets = spreadsheet_file.sheet_names
    # Template File
    wb = openpyxl.load_workbook(files['template_filename'])
    ws = wb['Projeto']
    # Skip 2 rows and start writing from row 3 - first two are headers in template file
    rownumber = 2

    for sheet_name in worksheets:
        df = pd.read_excel(spreadsheet_file, sheet_name)
        # Getting only the columns asked: "Part Number","QTY","Description","Material",
        # "Company","Category"
        df = df[["Part Number", "QTY", "Description", "Material", "Company", "Category"]]
        # Organizing info:
        # 1ยบ By Category
        # 2ยบ By Description
        df = df.sort_values(['Category', 'Description'],
                            ascending=[False, False])
        appended_data = df.to_dict()
        # Read all rows from df, but don't read index or header
        rows = dataframe_to_rows(df, index=False, header=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                # Write to cell, but after rownumber + row index
                ws.cell(row=r_idx + rownumber, column=c_idx, value=value)
        # Move the rownumber to end, so next worksheet data comes after this sheet's data 
        rownumber += len(df)
    wb.save('result.xlsx')


windows = Tk()
windows.withdraw()
windows.title("Column filter")
windows.geometry("500x150")

files = {}
get_excel_filename(files, 'input_filename', title='Select input file')
get_excel_filename(files, 'template_filename', title='Select template file')

windows.deiconify()
processing = Label(windows, text="Processing...", font=font.Font(size=30))
finished = Label(windows, text="Finished", font=font.Font(size=30))

processing.pack(fill='both')
execute_program(files)
processing.pack_forget()
finished.pack()

windows.mainloop()