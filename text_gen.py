from openpyxl import load_workbook
from docx import Document
from datetime import datetime
import shutil
import os
import re

def createComments():
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


def createSection(comments):
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


def retreiveDocumentInfo(file_path):
    try:
        # Load the Excel workbook
        wb = load_workbook(file_path)
        ws = wb['BM']
        print(f"Worksheet: {ws}")
        
        # Iterate over all cells with comments
        current_reserve_type = ws['B1'].value
        comments = createComments()
        
        document_text = []
        for col in ws.iter_cols():
            #verify if documment has finished & add the last type of reserve
            if col[1].column_letter == 'PD':  
                document_text.append(current_reserve_type)
                document_text.append(comments)
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
                        document_text.append(current_reserve_type)
                        document_text.append(comments)
                        current_reserve_type= col[0].value
                        comments= createComments()
                
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


def findBenchmark(field_path):
    #single file documentation
    directory_items= os.listdir(field_path)
    benchmarks= [item for item in directory_items if "Plantilla Benchmarking" in item]
    print('Error: there was no benchmark found(2)') if len(benchmarks) < 1 else None

    #find latest modified benchmark
    most_recent_date= datetime(1900, 1, 1)
    most_recent_file : str = None
    for file in benchmarks:
        new_file_date= datetime.fromtimestamp(os.path.getmtime(os.path.join(field_path, file)))
        if new_file_date > most_recent_date:
            most_recent_date= new_file_date
            most_recent_file= file
    print('Error: there was no benchmark found(1)') if not most_recent_file else None
    
    excel_path= os.path.join(field_path, most_recent_file)
    return excel_path


def traverseAsset(asset):
    #Get the folders in the directory
    asset_path= os.path.join(os.getcwd(), asset)
    print(asset)
    print(asset_path)
    asset_items= os.listdir(asset_path)
    fields= [item for item in asset_items if os.path.isdir(os.path.join(asset_path, item))]
    print(fields)

    #There's only one field in the asset
    if len(fields) < 1:
        excel_path= findBenchmark(asset_path)
        document_content = retreiveDocumentInfo(excel_path)
        print(excel_path)
        print(document_content)
    #there are multiple fields in the asset
    else:
        for field in fields:
            field_path= os.path.join(asset_path, field)
            excel_path= findBenchmark(field_path)
            document_content = retreiveDocumentInfo(excel_path)
            print(field)
            print(document_content)
            print()
            print()


  
traverseAsset('Cantagallo')
#traverseAsset('Cantagallo/Cristalina')