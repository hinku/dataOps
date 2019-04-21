'''
Created on 2019年4月13日

@author: Administrator
'''
import xlwings as xw
from dataOps import dictOps
from dataOps.updateOps import UpdateOps



def excelToDict(wb, sheetName, keyName, rowStartPos = 'A1'):
    '''
    将Excel指定的Sheet中的所有值转成双层dict保存
    如 keyName为a，数据获取从A1开始
    -----Excel内容----
    a b c
    1 2 3
    4 5 6

    返回结果为：
    {
        '1':{'a':1, 'b':2, 'c':3},
        '4':{'a':4, 'b':5, 'c':6},
    }

    '''
    sheet = wb.sheets[sheetName]
    datas = sheet.range(rowStartPos).expand()
    colNames = datas.value[0]
    print(colNames)
    keyIndexLst = [ i for i, v in enumerate(colNames) if v == keyName ]
    if keyIndexLst is None or len(keyIndexLst) == 0:
        print('has no col Named: %s' % keyName)
        return None
    keyIndex = keyIndexLst[0]
    dataDict = {}
    for data in datas.value[1:]:
        rowDict = { colNames[i]:v for i, v in enumerate(data) }
        dataDict[data[keyIndex]] = rowDict
    
    #print(dataDict)
    return dataDict

def getFormulaFromExcel(wb, sheetName, rowStartPos = 'A1'):
    '''
    获取指定sheet中各列的公式，如果不是公式时，该列公式记为None

    '''
    sheet = wb.sheets[sheetName]
    datas = sheet.range(rowStartPos).expand(mode='right')
    values = datas.value
    formulas = datas.formula

    retFormula = []
    for i, v in enumerate(values):
        #如果同个单元格的获取到的value和formula不一致，说明该单元格为公式，记录下公式，否则记录为None
        if v == formulas[i]:
            retFormula.append(None)
        else:
            retFormula.append(formulas[i])
    
    print(retFormula)
    return retFormula

class ExcelOps():
    '''
    classdocs
    '''
    app = None
    wb = None
    sheetName = None
    dataDict = None
    keyName = None
    dataDictValueTmp = None
    formulas = None
    def __init__(self, file, sheetName = 'sheet1', keyName = None):
        '''
        Constructor
        '''
        self.app = xw.App(visible=False,add_book=False)
        self.wb = self.app.books.open(file)
        self.sheetName = sheetName
        self.sheet = self.wb.sheets[self.sheetName]
        self.keyName = keyName
        self.dataDict = excelToDict(self.wb, self.sheetName, keyName)
        self.formulas = getFormulaFromExcel(self.wb, sheetName)
        #获取目的文件列格式
        for v in self.dataDict.values():
            self.dataDictValueTmp = v.copy()
            break
     
    def __del__(self):
        self.wb.save()
        self.wb.close()
        self.app.quit()
           
    def merge(self, *newFile, rowStartPos = 'A1'):
        for file in newFile:
            wb = self.app.books.open(file)
            newDataDict = excelToDict(wb, 0, self.keyName, rowStartPos)
            wb.close()
            for k, data in newDataDict.items():
                updateOps = UpdateOps.whileEmpty
                origData = self.dataDict.get(k)
                #如果文件中没有该条记录，需要将记录合并到文件，内容直接覆盖
                if origData is None:
                    origData = self.dataDictValueTmp
                    updateOps = UpdateOps.override
                    
                self.dataDict[k] = dictOps.merge(origData, data, updateOps)
            