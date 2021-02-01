import openpyxl
from pathlib import Path
import argparse
from typing import Optional, Sequence




def generate(sheet, exclude):
    # Generating column labels for output
    col_names = []
    for column in sheet.iter_cols(1, sheet.max_column):
        col_names.append(column[0].value)
    
    # Displaying the columns that are being excluded
    print(f'excluding columns: {[col_names[i] for i in range(0,len(col_names)) if i in exclude]}', end='\n\n')

    # formatted output logic
    row_num = 0
    for row in sheet.iter_rows():
        cell_num = 0
        if row_num != 0:
            for cell in row:
                if cell_num not in exclude and cell.value is not None:
                    print(f'{col_names[cell_num]}: {cell.value}', end='\n')
                cell_num += 1
            print('#' * 20)
        row_num += 1



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', help='input file name')
    parser.add_argument('-e','--exclude', nargs='+', help='columns to ignore')

    args = parser.parse_args()
    print(args)

    xlsx_file = Path('../test', 'test1.xlsx') 
    wb_obj = openpyxl.load_workbook(xlsx_file) 
    sheet = wb_obj.active
    
    exclude_nums = [int(i) for i in args.exclude]
    generate(sheet, set(exclude_nums))
    exit(0)

