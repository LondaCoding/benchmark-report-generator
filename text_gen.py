from openpyxl import load_workbook
import os
import re

def create_comments():
    arr = [['Variacion Reservas Netas'],
           ['Variacion Regalias ECP'],
           ['Variacion Revenue'],
           ['Variacion Fijo'],
           ['Variacion Semifijo'],
           ['Variacion OPEX Variable'],
           ['Variacion OPEX Ecopetrol'],
           ['Variacion Total CAPEX'],
           ['Variacion FC Operativo ECP'],
           ['Variacion Abandono']]
    return arr

def create_section(comments):
    text = ''

    #get the reserves that have comments
    reserves_with_comments= []
    reserves_without_comments= ''
    for reserve in comments:
        if len(reserve) > 1:
            reserves_with_comments.append(reserve)
        else:
            reserves_without_comments+= reserve[0]
            reserves_without_comments+= ', '

    for variable in reserves_with_comments:
        for comment in variable:
            text+= comment
        text+= '\n'
    
    #add the reserves_without_comments

    return text

def retrieve_comments(file_path):
    try:
        # Load the Excel workbook
        wb = load_workbook(file_path)
        ws = wb['BM']
        print(f"Worksheet: {ws}")
        
        # Iterate over all cells with comments
        current_reserve_type = ws['B1'].value
        comments = create_comments()
        
        document_text = ''
        for col in ws.iter_cols():
            #verify if documment has finished & add the last type of reserve
            if col[1].column_letter == 'PD':  
                document_text+= current_reserve_type + '\n'
                document_text+= create_section(comments)
                return document_text
            
            #verify if category has changed
            if col[0].value:
                if not(col[0].value==current_reserve_type) and not (col[0].value=="Variable"):
                    #print(col[0].value)
                    there_are_comments= False
                    for variable in comments:
                        if len(variable) > 1:
                            there_are_comments= True
                    if there_are_comments:
                        document_text+= '\n' + current_reserve_type + '\n'
                        document_text+= create_section(comments)
                        current_reserve_type= col[0].value
                        comments= create_comments()
                
            #traverse the column
            for cell in col:
                if cell.comment:
                    expression = r'Comment:\n(.*)' 
                    comment = f'{re.findall(expression, cell.comment.text, re.DOTALL)[0].replace(r'\n', '\n')}\n'
                    comments[(cell.row-3)//3].append(comment)
        return document_text        
                    
    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print(f"An error occurred: {e}")



# Concatenate the file name with the current directory using the os module
#file_path = os.path.join(os.path.dirname(__file__), file_name)
#c:\Users\juan.perez\OneDrive - Quorum Business Solutions\Por presentar a ECP\Aprobados ECP\Cantagallo\Cristalina\Plantilla Benchmarking - Campo Cristalina.xlsm
#c:\Users\juan.perez\OneDrive - Quorum Business Solutions\Documents\PS\Ecopetrol\Reservas\Automatization

#ile_path = str(input('File path: '))
#file_name = str(input('File name: '))
full_path = os.path.join('c:\\Users\\juan.perez\\OneDrive - Quorum Business Solutions\\Documents\\PS\\Ecopetrol\\Reservas\\Automatization', ('test1'+'.xlsm')).replace('\\', r'\\')
print(full_path)
document_text = retrieve_comments(full_path)
#print(document_text)
    