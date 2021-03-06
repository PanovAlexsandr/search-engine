
# find_word 1.2

### Table of Contents

1. [Installation](#installation)
2. [How to use](#htu)

## Installation <a name="installation"></a>
pip install find_word <br/>

## How to use <a name="htu"></a>
1. Add to your code:
```python 
from find_word import search_word
```  
<br/>

2. To start using all the features of the library you need to enter 2 commands:
```python 
a = search_word()
a.search('disk','format','word')
```  
<br/>
Where the:<br/>

```disk``` parameter is entered as the disk name ```('D', 'C', ...)```<br/>
```format```parameter is entered as the format in which the file is written:<br/>

```'txt'``` - text format<br/> 
```'doc'``` - word document format<br/> 
```'exl'``` - table file format exel<br/>
```'pdf'``` - Portable Document Format<br/>
```'all'``` - if you want to search in all formats at once<br/>
parameter ```word```:<br/>
```'word'``` - the search word<br/>

3. If you want to get a list of paths with all the files found, then use the construction:<br/>
```python 
a = search_word()
a.search('D', 'all','hello')
a.list()
```  
<br/>

4. If you want to open the found files use the construction:<br/>

```python 
a = search_word()
a.search('D', 'all','hello')
a.opening()
```  
By default, you can choose whether to open the file or not, but you can add any parameter and all files will open at once (example: ```a.opening('x')```
<br/>

5. If you want to copy the file, then you just need to enter the command:<br/>

```python 
a = search_word()
a.search('D', 'all','hello')
a.copying('Disk:\\path')
```  
Where in the function you write the path of copying files and add a new directory ```a.copying('D:\\copying_file')```<br/>
(you can only copy files to a new directory)
<br/>

## Example <a name="installation"></a>

```python
from find_word import search_word

a = search_word()
a.search('D', 'txt','word')
a.copying('E:\\d')
a.opening()
a.list() 
```
