import unittest, os
from find_enige import search_word

path = 'D:\\test.txt'
file = open(path, 'w')
file.write('hello world')
file.close()

class TestSearch_word(unittest.TestCase):
    a = search_word()
    a.search('D:\\', 'txt','hello')
    def test_list(self):
        b = self.a.list() 
        b = ''.join(b)
        self.assertEqual(b,path)

    def test_copying(self):
        self.a.copying('D:\\test')
        with open('D:\\test\\test.txt') as f:
            s = f.read()
            self.assertEqual(s,'hello world')
        
if __name__ == '__main__':
    unittest.main()



