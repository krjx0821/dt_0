# -*- coding: utf-8 -*-
"""
Created on Wed Oct 14 10:31:47 2020

@author: 19749
"""
import xlwings as xw
import docx
import comtypes.client
import os
path = 'C:\\Users/19749/Desktop/dt-main'
bat_name = '/w.bat'

class dt():
    def __init__(self, filename_list):
        self.l = filename_list
        
    def read_files(self):
        excel = []
        for title in self.l:
            if '.docx' in title:
                file=docx.Document(title)
                tables = file.tables #获取文件中的表格集
                table_list = []
                for table in tables[:]:
                    table_content = []
                    for i, row in enumerate(table.rows[:]):   # 读每行
                        row_content = []
                        for cell in row.cells[:]:  # 读一行中的所有单元格
                            c = cell.text
                            row_content.append(c)
                        table_content.append(row_content)
                    table_list.append(table_content)
                
                for table in table_list:
                    flag = 0
                    for row in table:
                        if row[0] == '阶段':
                            flag = 1
                    if flag == 1:
                        excel.append([title])
                        excel += table
                        excel.append([])
                print(excel)
  
        sht = xw.Book().sheets('sheet1')  # 新增一个表格
        for i in range(len(excel)):
            print('len(excel):', len(excel))
            print('i:', i)
            sht.range('A{0}'.format(str(i+1))).value = excel[i]

def test_excel_macro(path):
    App = comtypes.client.CreateObject('Word.Application')
    print(2)
    App.Documents.Open(path+'/doc2docx.doc')
    print(3)
    App.Application.Run('doc2docx')
    print(4)
    App.Documents(1).Close(SaveChanges=0)
    print(5)
    App.Application.Quit()
    print(6)  

if __name__ == "__main__":
    test_excel_macro(path)
    print(7)
    os.system(path + bat_name)
    print(8)
    f = open('LIST.TXT')
    print(9)
    a = f.read()
    print(10)
    filename_list = a.split('\n')
    dt = dt(filename_list)
    dt.read_files()
    file = open('LIST.txt', 'w').close()                        