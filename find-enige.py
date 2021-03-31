import os
import subprocess
import docx
import openpyxl
import PyPDF2


class search_word:

    spisok_file1 = []
    spisok_file2 = []
    spisok_file3 = []
    spisok_file4 = []

    find_file1 = []
    find_file2 = []
    find_file3 = []
    find_file4 = []

    def search(self, disk, format_files, word):
        doc = ('.doc', '.docx')
        exl = ('.xls', '.xlsx')
        t = 'txt'
        d = 'doc'
        e = 'exl'        
        p = 'pdf'
        a = 'all'
        
        #search 
        if format_files == t or format_files == a: 
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    s = (os.path.join(adress, file))
                    if file.endswith('.txt') and '$' not in s:
                        self.spisok_file1.append(s)

            #search by word
            for x in self.spisok_file1: 
                with open(x) as r:
                    try:
                        for line in r:
                            if word in line:   
                                self.find_file1.append(x)  
                    except Exception as fail:
                        with open('fail_file.txt', 'w') as r:
                            r.write(str(fail)+'\n')  
        
            #open file 
            
            
            for x in self.find_file1:
                file_name = x                    
                print('We faind faile, her path: ', file_name)
                want_open = input('Open file ? \n (y/n): ')
                if want_open == 'y':
                    directory = file_name
                    subprocess.Popen('explorer ' + directory)
                
      

        if format_files == d or format_files == a: 
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    s = (os.path.join(adress, file))
                    if file.endswith(doc) and '$' not in s:
                        self.spisok_file2.append(s)
            
            for x in self.spisok_file2: 
                with open(x) as r:
                    try:
                        doc = docx.Document(x)
                        text = []
                        for paragraph in doc.paragraphs:
                            text.append(paragraph.text)    
                        for line in text:
                            if word in line:                        
                                self.find_file2.append(x)

                    except Exception as fail:
                        with open('fail_file.txt', 'w') as r:
                            r.write(str(fail)+'\n')  
                
            for x in self.find_file2:
                file_name = x
                print('We faind faile, her path: ', file_name)
                want_open = input('Open file ? \n (y/n): ')
                if want_open == 'y':
                    directory = file_name
                    subprocess.Popen('explorer ' + directory)

        if format_files == e or format_files == a: 
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    s = (os.path.join(adress, file))
                    if file.endswith(exl) and '$' not in s:
                        self.spisok_file3.append(s)

            for x in self.spisok_file3: 
                with open(x) as r:
                    try:
                        path = x
                        wb_obj = openpyxl.load_workbook(path)
                        sheet_obj = wb_obj.active
                        max_col = sheet_obj.max_column
                        for i in range(1, max_col + 1):
                            cell_obj = sheet_obj.cell(row = 1, column = i)
                            if word in cell_obj.value:                            
                                self.find_file3.append(x)
                    except Exception as fail:
                        with open('fail_file.txt', 'w') as r:
                            r.write(str(fail)+'\n') 

            for x in self.find_file3:
                file_name = x
                print('We faind faile, her path: ', file_name)
                want_open = input('Open file ? \n (y/n): ')
                if want_open == 'y':
                    directory = file_name
                    subprocess.Popen('explorer ' + directory)

        if format_files == p or format_files == a: 
            for adress, dirs, files in os.walk(disk):
                for file in files:
                    s = (os.path.join(adress, file))
                    if file.endswith('.pdf') and '$' not in s:
                        self.spisok_file4.append(s)
            for x in self.spisok_file4: 
                with open(x) as r:
                    try:
                        pdf_file = open(x, 'rb')
                        read_pdf = PyPDF2.PdfFileReader(pdf_file)
                        page = read_pdf.getPage(0)
                        page_content = page.extractText()
                        if word in page_content:                            
                            self.find_file4.append(x)
                        pdf_file.close()

                    except Exception as fail:
                        with open('fail_file.txt', 'w') as r:
                            r.write(str(fail)+'\n')  

            for x in self.find_file4:
                file_name = x
                print('We faind faile, her path: ', file_name)
                want_open = input('Open file ? \n (y/n): ')
                if want_open == 'y':
                    directory = file_name
                    subprocess.Popen('explorer ' + directory)

        #delete fail
        if os.path.exists("fail_file.txt") == True:
            os.remove('fail_file.txt')
search_word().search('D:\\', 'all', 'hel')