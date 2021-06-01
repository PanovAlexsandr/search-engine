
# Find enige 1.1

### Table of Contents

1. [Installation](#installation)
2. [How to use](#htu)

## Installation <a name="installation"></a>
pip install find_enige <br/>

## How to use <a name="htu"></a>
1. Add to your code:
```python 
from DoubleLinkedList import DLinked
```  
<br/>

2. To start using all the features of the library you need to enter 2 commands:
```python 
a = search_word()
a.search('disk','format','word')
```  
<br/>
Where the:<br/>
disk parameter is entered as the disk name (```'D', 'C', ...```)<br/>
format parameter is entered as the format in which the file is written:<br/>
```'txt'``` - text format<br/> 
```'doc'``` - word document format<br/> 
```'exl'``` - table file format exel<br/>
```'pdf'``` - Portable Document Format<br/>
```'all'``` - if you want to search in all formats at once<br/>
```'word'``` - the search word<br/>

3. If you want to get a list of paths with all the files found, then use the construction:<br/>
```python 
a = search_word()
a.search('D', 'all','hello')
a.list()
```  
<br/>
If you want to open the found files use the construction:<br/>

```python 
a = search_word()
a.search('D', 'all','hello')
a.opening()
```  
<br/>
By default, you can choose whether to open the file or not, but you can add any parameter and all files will open at once (example: ```a.opening(' ')```
<br/>

4. If you want to copy the file, then you just need to enter the command:<br/>

```python 
a = search_word()
a.search('D', 'all','hello')
a.copying('Disk:\\path')
```  


## Example <a name="installation"></a>

```python
a = search_word()
a.search('D', 'all','hello')
a.copying('E:\\d')
a.opening()
a.list() 
```
