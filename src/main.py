import openpyxl
from pathlib import Path
import argparse
from typing import Optional, Sequence




def generate(sheet):
    exclude = 3 
    print(f'excluding: {exclude}')
    col_names = []
    for column in sheet.iter_cols(1, sheet.max_column):
        col_names.append(column[0].value)
    
    print(col_names, end='\n\n')
    
    row_num = 0
    for row in sheet.iter_rows():
        cell_num = 0
        if row_num != 0:
            for cell in row:
                if cell_num != exclude:
                    print(f'{col_names[cell_num]}: {cell.value}', end='\n')
                cell_num += 1
            print('#' * 20)
        row_num += 1




if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', help='input file name')
    args = parser.parse_args()
    print(args)

    xlsx_file = Path('../test', 'test1.xlsx') 
    wb_obj = openpyxl.load_workbook(xlsx_file) 
    sheet = wb_obj.active
    
    generate(sheet)
    exit(0)

