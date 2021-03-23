import os
import subprocess
import docx
import openpyxl
import PyPDF2
import re

flag = True
while flag:
    disk = input('Disk name : ')
    word = input('Input word : ')
    accuracy = input('Do you want an exact match or not ?  \n (y/n): ')
    spisok = []
    num = 0 
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
                            rgx = re.compile("(\w[\w']*\w|\w)")
                            out=rgx.findall(line)
                            if accuracy == 'n':
                                if word in line:                            
                                    faind_file.write(sf + '\n')
                                    num = 1
                            elif accuracy == 'y':
                                a = ' '
                                if a + word + a in line:                            
                                        faind_file.write(sf + '\n')
                                        num = 1
                                elif out[0] == word:
                                    faind_file.write(sf + '\n')
                                    num = 1
                                elif out[-1] == word:
                                    faind_file.write(sf + '\n')
                                    num = 1
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
                            for line in text:
                                rgx = re.compile("(\w[\w']*\w|\w)")
                                out=rgx.findall(line)
                        if accuracy == 'n':
                            if word in text:                            
                                faind_file.write(sf + '\n')
                                num = 1
                        elif accuracy == 'y':
                            a = ' '
                            if a + word + a in text:                            
                                faind_file.write(sf + '\n')
                                num = 1  
                            elif out[0] == word:
                                    faind_file.write(sf + '\n')
                                    num = 1
                            elif out[-1] == word:
                                faind_file.write(sf + '\n')
                                num = 1  
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
                            rgx = re.compile("(\w[\w']*\w|\w)")
                            out=rgx.findall(cell_obj.value)
                            if accuracy == 'n':
                                if word in cell_obj.value:                            
                                    faind_file.write(sf + '\n')
                                    num = 1
                            elif accuracy == 'y':
                                a = ' '
                                if a + word + a in cell_obj.value:                            
                                        faind_file.write(sf + '\n')
                                        num = 1
                                elif out[0] == word:
                                    faind_file.write(sf + '\n')
                                    num = 1
                                elif out[-1] == word:
                                    faind_file.write(sf + '\n')
                                    num = 1  
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
                        rgx = re.compile("(\w[\w']*\w|\w)")
                        out=rgx.findall(page_content)
                        if accuracy == 'n':
                            if word in page_content:                            
                                faind_file.write(sf + '\n')
                                num = 1
                        elif accuracy == 'y':
                            a = ' '
                            if a + word + a in page_content:                            
                                faind_file.write(sf + '\n')
                                num = 1
                            elif out[0] == word:
                                faind_file.write(sf + '\n')
                                num = 1
                            elif out[-1] == word:
                                faind_file.write(sf + '\n')
                                num = 1
                        pdf_file.close()


                except Exception as fail:
                    with open('fail_file.txt', 'a+') as f:
                        f.write(str(fail)+'\n')
                finally:
                    f.close()
                    

    #open file
    
    if num == 1:
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

    else: print('We not find file')

    flag = True if input(' start over? \n (y/n): ') == 'y' else False 
os.remove('search_files.txt')
os.remove('faind_file.txt')
if os.path.exists("fail_file.txt") == True:
    os.remove('fail_file.txt')