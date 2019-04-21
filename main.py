'''
Created on 2019年4月13日

@author: Administrator
'''

from dataOps.excelOps import ExcelOps
import os
if __name__ == '__main__':
    #path = os.path.join(os.getcwd(), 'dataOps')
    path = os.getcwd()

    xls = ExcelOps(os.path.join(path, 'test1.xlsx'), keyName = 'a')
    print(xls.dataDict)
    xls.merge(os.path.join(path, 'test2.xlsx'), rowStartPos = 'A2')
    print(xls.dataDict)