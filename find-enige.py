import os
import subprocess
import docx
import openpyxl
import PyPDF2

flag = True
while flag:
    disk = input('Disk name : ')
    word = input('Input word : ')
    spisok = []
    t = ['1 = txt','2 = doc','3 = xls','4 = pdf','5 = all formats']
    print(*t, sep = '\n')
    choice = int(input('Select a formats: '))
    doc = ('.doc', '.docx')
    exl = ('.xls', '.xlsx')
    full = ('.txt', '.pdf') + doc + exl

    #search
    if choice == 1: 
        for adress, dirs, files in os.walk(disk):
            for file in files:
                s = (os.path.join(adress, file))
                if file.endswith('.txt') and '$' not in s:
                    spisok.append(s)
    elif choice == 2:
        for adress, dirs, files in os.walk(disk):
            for file in files:
                s = (os.path.join(adress, file))
                if file.endswith(doc) and '$' not in s:
                    spisok.append(s)
    elif choice == 3:
        for adress, dirs, files in os.walk(disk):
            for file in files:
                s = (os.path.join(adress, file))
                if file.endswith(exl) and '$' not in s:
                    spisok.append(s)
    elif choice == 4:
        for adress, dirs, files in os.walk(disk):
            for file in files:
                s = (os.path.join(adress, file))
                if file.endswith('.pdf')  and '$' not in s:
                    spisok.append(s)
    else:
        for adress, dirs, files in os.walk(disk):
            for file in files:
                s = (os.path.join(adress, file))
                if file.endswith(full) and '$' not in s:
                    spisok.append(s)

                    
    #recording found files
    r = open('search_files.txt', 'w')
    for x in spisok:
        r.write(x + '\n')
    r.close()
    


    #search by word
    with open('search_files.txt') as r:
        with open('faind_file.txt', 'w') as faind_file:
            for line in r:
                sf = line[0:-1]
                f = open(sf) 
                try:
                    if choice == 1 or choice == 5:
                        for line in f:
                            if word in line:                            
                                faind_file.write(sf + '\n')
                except Exception as fail:
                    with open('fail_file.txt', 'w') as f:
                        f.write(str(fail)+'\n')
                finally:
                    f.close()

                f = open(sf) 
                try:    
                    if choice == 2 or choice == 5:
                        doc = docx.Document(sf)
                        text = []
                        for paragraph in doc.paragraphs:
                            text.append(paragraph.text)
                        if word in text:                            
                            faind_file.write(sf + '\n')
                except Exception as fail:
                    with open('fail_file.txt', 'a+') as f:
                        f.write(str(fail)+'\n')
                finally:
                    f.close()

                f = open(sf) 
                try:      
                    if choice == 3 or choice == 5:
                        path = sf
                        wb_obj = openpyxl.load_workbook(path)
                        sheet_obj = wb_obj.active
                        max_col = sheet_obj.max_column
                        for i in range(1, max_col + 1):
                            cell_obj = sheet_obj.cell(row = 1, column = i)
                            print(cell_obj.value)
                            if word in cell_obj.value:                            
                                faind_file.write(sf + '\n')
                except Exception as fail:
                    with open('fail_file.txt', 'a+') as f:
                        f.write(str(fail)+'\n')
                finally:
                    f.close()
                
                f = open(sf) 
                try:    
                    if choice == 4 or choice == 5:
                        pdf_file = open(sf, 'rb')
                        read_pdf = PyPDF2.PdfFileReader(pdf_file)
                        #number_of_pages = read_pdf.getNumPages()
                        page = read_pdf.getPage(0)
                        page_content = page.extractText()
                        if word in page_content:                            
                            faind_file.write(sf + '\n')
                        pdf_file.close()


                except Exception as fail:
                    with open('fail_file.txt', 'a+') as f:
                        f.write(str(fail)+'\n')
                finally:
                    f.close()
                    

    #open file
    if os.path.exists("faind_file.txt") == True:
        faind_file = open('faind_file.txt')

        try:
            for line in faind_file:
                count = 1 + sum(1 for line in faind_file)   
                if count == 1:
                    print('We find file')
                    want_open = input('Open file ? \n (y/n): ')
                    if want_open == 'y':
                        new_line = line.replace("\\\\","\\")
                        directory = new_line
                        subprocess.Popen('explorer ' + directory)
                elif count > 1:
                    print('We find',count,'files')
                    want_open = input('Open files ? \n (y/n): ')
                    if want_open == 'y':
                        with open('faind_file.txt') as faind_file:
                            for line in faind_file:    
                                new_line = line.replace("\\\\","\\")
                                directory = new_line
                                subprocess.Popen('explorer ' + directory)
        finally:
            faind_file.close()
    flag = True if input(' start over? \n (y/n): ') == 'y' else False 
os.remove('search_files.txt')
os.remove('faind_file.txt')
if os.path.exists("fail_file.txt") == True:
    os.remove('fail_file.txt')