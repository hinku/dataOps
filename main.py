'''
Created on 2019年4月13日

@author: Administrator
'''

from dataOps.excelOps import ExcelOps
if __name__ == '__main__':
    xls = ExcelOps(r'test1.xlsx', keyName = 'a')
    print(xls.dataDict)
    xls.merge(r'test2.xlsx', rowStartPos = 'A2')
    print(xls.dataDict)