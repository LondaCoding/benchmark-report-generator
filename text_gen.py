from openpyxl import load_workbook
import os
import re

def retrieve_comments(file_path):
    try:
        # Load the Excel workbook
        wb = load_workbook(file_path)

        # Iterate over all worksheets
        ws = wb['BM']
        print(f"Worksheet: {ws}")
        
        # Iterate over all cells with comments
        col_change = ws['B1'].value
        print('first:', col_change)
        
        for col in ws.iter_cols():
            if not (col[0].value == col_change):
                col[0]
                col_change = col[0].value
                print_col = True 
            print_col = False
                
            #print the column
            for cell in col:
                if cell.comment:
                    if print_col:
                        print('Col iterations:', col_change)
                    #print the variable
                    variable_loc = 'A' + str( int(cell.row) + ((int(cell.row)-2) % 3) - 1) 
                    print(ws[variable_loc].value, end=None)

                    expression = r'Comment:\n(.*)' 
                    print(f'{re.findall(expression, cell.comment.text, re.DOTALL)[0].replace(r'\n', '\n')}\n')
                    #print(f"Cell {cell.coordinate}: {re.findall(expression, cell.comment.text, re.DOTALL)[0].replace(r'\n', '\n')}\n")
                    
                    
    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print(f"An error occurred: {e}")



# Concatenate the file name with the current directory using the os module
#file_path = os.path.join(os.path.dirname(__file__), file_name)
#C:\Users\juan.perez\OneDrive - Quorum Business Solutions\Por presentar a ECP\Aprobados ECP\Cantagallo\Cristalina\Plantilla Benchmarking - Campo Cristalina.xlsm
file_path = str(input('File path: '))
file_name = str(input('File name: '))
full_path = os.path.join(file_path, (file_name+'.xlsm')).replace('\\', r'\\')
print(full_path)
retrieve_comments(full_path)
    