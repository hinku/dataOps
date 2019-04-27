'''
Created on 2019年4月13日

@author: Administrator
'''
import xlwings as xw
from .updateOps import UpdateOps
from . import dictOps




def excelToDict(wb, sheetName, keyName, startCell = 'A1'):
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
    datas = sheet.range(startCell).expand()
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

def getFormulaFromExcel(wb, sheetName, startCell, endCell):
    '''
    获取指定sheet中各列的公式，如果不是公式时，该列公式记为None

    '''
    sheet = wb.sheets[sheetName]
    datas = sheet.range(startCell, endCell)
    values = datas.value
    formulas = datas.formula[0]
    
    retFormula = []
    for i, v in enumerate(values):
        #如果同个单元格的获取到的value和formula不一致，说明该单元格为公式，记录下公式，否则记录为None
        if v != formulas[i] and formulas[i].startswith('='):
            retFormula.append(formulas[i])
        else:
            retFormula.append(None)
    
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
    __emptyColumnDict = None
    formulas = None
    startCell = None

    def getEmptyColumn(self):
        '''
            获取构造一个空列数据
        '''
        for v in self.dataDict.values():
            return { k: None for k in v }

    def __init__(self, file, sheetName = 'sheet1', keyName = None, startCell = 'A1'):
        '''
        Constructor
        '''
        self.app = xw.App(visible=False,add_book=False)
        self.wb = self.app.books.open(file)
        self.sheetName = sheetName
        self.sheet = self.wb.sheets[self.sheetName]
        self.keyName = keyName
        self.dataDict = excelToDict(self.wb, self.sheetName, keyName)
        self.startCell = self.sheet.range(startCell)
        self.formulas = getFormulaFromExcel(self.wb, sheetName, startCell, (self.startCell.row, len(self.getEmptyColumn().keys())))
        
        #获取目的文件列格式
        #self.emptyColumnDict = self.getEmptyColumn()

     
    #析构操作
    def __del__(self):
        '''
        析构函数操作：保存文件，退出wb、app
        '''
        self.wb.save()
        self.wb.close()
        self.app.quit()
    
    #将数据写入excel
    def flush(self):
        '''
        将数据写入excel
        '''
        values = []
        for v in self.dataDict.values():
            #print(v)
            values.append(list(v.values()))
        
        valueStartRow = self.startCell.row + 1
        valueEndRow = valueStartRow + len(values) - 1
        self.sheet.range('A' + str(valueStartRow)).value = values
        
        #对于公式部分，需要从新对单元列赋值
        print(self.formulas)
        for i, formula in enumerate(self.formulas):
            if formula is not None:
                self.sheet.range((valueStartRow, i+1), (valueEndRow, i+1)).formula = formula

        self.wb.save()
        
    
    #将
    def merge(self, *newFile, sheet = 0, startCell = 'A1'):
        '''
        将新制定的文件，将源文件有的且新文件也有的字段更新到源文件中；如果新文件中的记录在源文件中不存在，追加进去
        '''
        for file in newFile:
            wb = self.app.books.open(file)
            newDataDict = excelToDict(wb, 0, self.keyName, startCell)
            wb.close()
            for k, newData in newDataDict.items():
                updateOps = UpdateOps.whileEmpty
                origData = self.dataDict.get(k)
                #如果文件中没有该条记录，需要将记录合并到文件，内容直接覆盖
                if origData is None:
                    origData = self.getEmptyColumn()
                    updateOps = UpdateOps.override
                    
                self.dataDict[k] = dictOps.merge(origData, newData, updateOps)
            
