from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, RGBColor
from datetime import datetime
import sys
import time
import os
import re

error_log= []

def createComments():
    arr = [['Reservas Netas'],
           ['Regalias'],
           ['Revenue'],
           ['OPEX Fijo'],
           ['OPEX Semifijo'],
           ['OPEX Variable'],
           ['OPEX Ecopetrol'],
           ['CAPEX Total'],
           ['Flujo de Caja Operativo'],
           ['Abandono']]
    return arr


def retreiveDocumentInfo(file_path, worksheet):
    try:
        # Load the Excel workbook
        wb = load_workbook(file_path)
        ws = wb[worksheet]
        print(f"Retreiving comments from Worksheet: {ws}...")
    except FileNotFoundError as e:
        return str(e)
    except Exception as e:
        return str(e)

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
                
                expression = re.compile(r'Comment:\n(.*)', re.DOTALL)
                comment = expression.search(cell.comment.text)
                if not comment:
                    expression = re.compile(r'Comentario:\n(.*)', re.DOTALL)
                    comment = expression.search(cell.comment.text)
                comment= comment.group(1)
                comments[(cell.row-3)//3].append(comment)
    return document_text        


def findBenchmark(field_path):
    #single file documentation
    directory_items= os.listdir(field_path)
    benchmarks= [item for item in directory_items if "Plantilla Benchmarking" in item]
    if len(benchmarks) < 1:
        error_message= 'Error: there was no benchmark found (1).'
        return error_message
    else: 
        None

    #find latest modified benchmark
    most_recent_date= datetime(1900, 1, 1)
    most_recent_file : str = None
    for file in benchmarks:
        new_file_date= datetime.fromtimestamp(os.path.getmtime(os.path.join(field_path, file)))
        if new_file_date > most_recent_date:
            most_recent_date= new_file_date
            most_recent_file= file

    if not most_recent_file:
        error_message= 'Error: there was no benchmark found (2)'
        return error_message
    else:
        excel_path= os.path.join(field_path, most_recent_file)
        return excel_path


def addFieldToDocument(doc:Document, field_info, field_name):
    accepted_types= {'PDP':'Desarrolladas produciendo', 
               'PNP':'Desarrolladas no produciendo', 
               'PND':'No desarrolladas',
               'PRB':'Probables',
               'PS':'Posibles'}
    
    doc.add_heading(field_name, level=1)
    print_type= True
    for location_type in field_info:
        #skip if not an accepted type type
        if print_type==False:
            print_type=True
            continue
        #it's a header 1
        if type(location_type) is str:
            if location_type in accepted_types:
                title_text= accepted_types[location_type]+f' ({location_type})'
                doc.add_heading(title_text, level=2)
            else:
                print_type= False
                print(f'The type: "{location_type}" was skipped')
        #it's a variable array
        else:
            #Get variables that don't havecomments
            no_comment_variables= []
            for variable in location_type:
                if len(variable) <= 1:
                    no_comment_variables.append(variable[0])
                    location_type.remove(variable)
            no_comment_variables_title : str = ''
            for variable in no_comment_variables:
                no_comment_variables_title+= variable + ', '
            no_comment_variables_title= no_comment_variables_title[:-2]
            doc.add_heading(no_comment_variables_title, level=3)
            doc.add_paragraph('Calculo OK')

            #add comments
            for variable in location_type:
                if len(variable) > 1:
                    doc.add_heading(variable[0], level=3)
                    for comment in variable[1:]:
                        comment= comment.strip(' \t\n\r')
                        doc.add_paragraph(comment)
            doc.add_paragraph('')            
    doc.add_paragraph('')           

            
    return doc


def traverseAsset(asset, worksheet):
    template_path= os.path.join(os.getcwd(), "Template.docx") #change to Template!!!!!!!!!!
    doc= Document(template_path)
    
    #Get the folders in the directory
    asset_path= os.path.join(os.getcwd(), asset)
    asset_items= os.listdir(asset_path)
    fields= [item for item in asset_items if os.path.isdir(os.path.join(asset_path, item))]

    #modify the title & it's style
    regex= r'\d+. (.*)'
    try:
        asset_name= re.findall(regex, asset)[0]
    except IndexError as e:
        print(f'Error traversing the asset, the folder numeration <#. > finished, there are no more finished assets.')
        print('Program finished')
        sys.exit(0)

    print('Name of the asset:', asset_name)

    doc.paragraphs[16].text= asset_name
    target_paragraph = doc.paragraphs[16] 
    for run in target_paragraph.runs:
        run.font.size = Pt(26)  # Change to the desired font size in points
    for run in target_paragraph.runs:
        run.font.color.rgb = RGBColor(0, 112, 192)  # Change RGB values for the desired color
    for run in target_paragraph.runs:
        run.font.name = 'Calibri'  # Change to the desired font name

    #There's only one field in the asset
    if len(fields) < 1:
        print("Theres only one field in the asset")
        temporal_doc= createField(doc, asset_path, asset_name, asset_name, worksheet)
        if temporal_doc:
            doc=temporal_doc 
        else:
            return None
    #there are multiple fields in the asset
    else:
        print(f'Fields of the asset "{asset}":', fields)
        for field in fields:
            field_path= os.path.join(asset_path, field)
            temporal_doc= createField(doc, field_path, field, asset_name, worksheet)
            if temporal_doc:
                doc=temporal_doc 
            else:
                print('The field info was not added to the document')
                continue

    #Save file
    new_file_location= os.path.join(os.getcwd(), f'Reportes Generados\\{asset}.docx')
    doc.save(new_file_location)
    print(f"Document {new_file_location} created")

  
def createField(doc, field_path, field, asset, worksheet):
    print()
    print('Creating field:', field)
    print('Finding benchmark...')
    excel_path= findBenchmark(field_path)
    if 'Error:' in excel_path:
        error_message= f'Couldnt find the benchmark for the only field of "{field}" of asset: "{asset}": {excel_path}'
        error_log.append(error_message)
        print(error_message)
        return None
    
    print(f'Fetching field "{field}" comments...')
    excel_content= retreiveDocumentInfo(excel_path, worksheet)
    if type(excel_content) is str:
        error_message= f'Error retreiving field "{field}" comments of asset: "{asset}": {excel_content}'
        error_log.append(error_message)
        print(error_message)
        return 
    
    print(f'Adding field "{field}" comments...')
    return addFieldToDocument(doc, excel_content, field)


def generateReportFolder():
    #create directory where reports will be saved
    if not "Reportes Generados" in os.listdir(os.getcwd()):
        print('Creating directory "Reportes Generados"...')
        os.makedirs("Reportes Generados") 
    else:
        print('Directory "Reportes Generados" already exists')
        try:
            for file in os.listdir("Reportes Generados"):
                os.remove(f"Reportes Generados\\{file}")
                print(f'The file "Reportes Generados\\{file}" has been erased')
            time.sleep(1)
        except Exception as e:
            error_message= f"An error occurred: {e}\n----------------RE-RUN THE PROGRAM AFTER CLOSING ALL DOCUMMENTS TO ENSURE INTEGRITY----------------"
            error_log.append(error_message)
            print(error_message)


    assets= [item for item in os.listdir(os.getcwd()) if os.path.isdir(item)]
    try:
        assets.remove("Reportes Generados")
    except ValueError:
        None

    print('Assets:', assets)
    print()
    worksheet= str(input("Ingrese el nombre del WorkSheet de los documentos: "))
    for asset in assets:
        print('CREATING DOCUMENT:', asset)
        traverseAsset(asset, worksheet)
        print()
        print()

    print('ERRORS DURING EXECUTION')
    for error in error_log:
        print(error)
        print()


generateReportFolder()
