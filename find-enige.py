import os

disk = input('Disk name : ')
word = input('Input word : ')
spisok = []



#поиск файлов формата txt
for adress, dirs, files in os.walk(disk):
   for file in files:
       s = (os.path.join(adress, file))
       if file.endswith('.txt') and '$' not in s:
           spisok.append(s)



#запись всех их в список
r = open('search_files.txt', 'w')
for x in spisok:
    r.write(x + '\n')
r.close()



#поиск слова
with open('search_files.txt') as r:
    for line in r:
        sf = line[0:-1]
        f = open(sf)
        try:
            for line in f:
                if word in line:
                    print(f)
                    print(line, end ='' + '\n')
        finally:
            f.close()


#запись файлов не открывшихся в исключение
  

