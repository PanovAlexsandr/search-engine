import os
import subprocess
import docx
import openpyxl
import PyPDF2
import re
import pathlib
flag = True
while flag:
    disk = input('Disk name : ')
    word = input('Input word : ')
    accuracy = input('Do you want an exact match or not ?  \n (y/n): ')
    t = ['1 = txt','2 = doc','3 = xls','4 = pdf']
    print(*t, sep = '\n')
    choice = int(input('Select a formats: '))
    doc = ('.doc', '.docx')
    exl = ('.xls', '.xlsx')

    #search
    def search():
        if choice == 1: 
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    if file.endswith('.txt') and '$' not in file:
                        yield os.path.join(adress, file)
        elif choice == 2:
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    if file.endswith(doc) and '$' not in file:
                        yield os.path.join(adress, file)
        elif choice == 3:
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    if file.endswith(exl) and '$' not in file:
                        yield os.path.join(adress, file)
        elif choice == 4:
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    if file.endswith('.pdf')  and '$' not in file:
                        yield os.path.join(adress, file)

    #read file and search by word
    def read_file(path):
        with open(path) as r:
            if choice == 1:
                for line in r:
                    if accuracy == 'n':
                        if word in line:                           
                            return open_file(path)
                    elif accuracy == 'y':
                        rgx = re.compile("(\w[\w']*\w|\w)")
                        out=rgx.findall(line)
                        a = ' '
                        if a + word + a in line:              
                            return open_file(path)
                        elif out[0] == word:
                            return open_file(path)
                        elif out[-1] == word:
                            return open_file(path)
        
            elif choice == 2:
                print(1)   
                doc = docx.Document(path)
                text = []
                for paragraph in doc.paragraphs:
                    text.append(paragraph.text)
                    for line in text:
                        rgx = re.compile("(\w[\w']*\w|\w)")
                        out=rgx.findall(line)
                if accuracy == 'n':
                    if word in text:                            
                        return open_file(path) 
                elif accuracy == 'y':
                    a = ' '
                    if a + word + a in text:                            
                        return open_file(path)
                    elif out[0] == word:
                        return open_file(path)
                    elif out[-1] == word:
                        return open_file(path)

            elif choice == 3:
                wb_obj = openpyxl.load_workbook(path)
                sheet_obj = wb_obj.active
                max_col = sheet_obj.max_column
                for i in range(1, max_col + 1):
                    cell_obj = sheet_obj.cell(row = 1, column = i)
                    rgx = re.compile("(\w[\w']*\w|\w)")
                    out=rgx.findall(cell_obj.value)
                    if accuracy == 'n':
                        if word in cell_obj.value:                            
                            return open_file(path)
                    elif accuracy == 'y':
                        a = ' '
                        if a + word + a in cell_obj.value:                            
                            return open_file(path)
                        elif out[0] == word:
                            return open_file(path)
                        elif out[-1] == word:
                            return open_file(path)

            elif choice == 4:
                pdf_file = open(path, 'rb')
                read_pdf = PyPDF2.PdfFileReader(pdf_file)
                page = read_pdf.getPage(0)
                page_content = page.extractText()
                rgx = re.compile("(\w[\w']*\w|\w)")
                out=rgx.findall(page_content)
                if accuracy == 'n':
                    if word in page_content:                            
                        return open_file(path)
                elif accuracy == 'y':
                    a = ' '
                    if a + word + a in page_content:                            
                        return open_file(path)
                    elif out[0] == word:
                        return open_file(path)
                    elif out[-1] == word:
                        return open_file(path)
                pdf_file.close()        

    #open file
    def open_file(path):
        file_name = path.split('\\')[-1]
        print('We faind faile:', file_name)
        want_open = input('Open file ? \n (y/n): ')
        if want_open == 'y':
            new_line = path.replace("\\\\","\\")
            directory = new_line
            subprocess.Popen('explorer ' + directory)


    for line in search():
        try:
            read_file(line)
        except Exception as fail:
                    with open('fail_file.txt', 'w') as r:
                        r.write(str(fail)+'\n')
    
    flag = True if input(' start over? \n (y/n): ') == 'y' else False 
if os.path.exists("fail_file.txt") == True:
    os.remove('fail_file.txt')