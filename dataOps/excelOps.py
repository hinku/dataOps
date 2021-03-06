'''
Created on 2019年4月13日

@author: Administrator
'''
import xlwings as xw
from .updateOps import UpdateOps
from . import dictOps




def excelToDict(wb, sheetNameOrIndex, keyName, startCell = 'A1'):
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
    sheet = wb.sheets[sheetNameOrIndex]
    datas = sheet.range(startCell).expand()
    colNames = datas.value[0]
    print(colNames)
    keyIndexLst = [ i for i, v in enumerate(colNames) if v == keyName ]
    if not keyIndexLst:
        print('has no col Named: %s' % keyName)
        return None
    keyIndex = keyIndexLst[0]
    dataDict = {}
    for data in datas.value[1:]:
        rowDict = { colNames[i]:v for i, v in enumerate(data) }
        dataDict[data[keyIndex]] = rowDict
    
    #print(dataDict)
    return dataDict

def getFormulaFromExcel(wb, sheetNameOrIndex, startCell, endCell):
    '''
    获取指定sheet中各列的公式，如果不是公式时，该列公式记为None

    '''
    sheet = wb.sheets[sheetNameOrIndex]
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
    sheetNameOrIndex = None
    dataDict = None
    keyName = None
    formulas = None
    startCell = None

    def getEmptyColumn(self):
        '''
            获取构造一个空列数据
        '''
        for v in self.dataDict.values():
            return { k: None for k in v }

    def __init__(self, file, sheetNameOrIndex = 0, keyName = None, startCell = 'A1'):
        '''
        Constructor
        '''
        self.app = xw.App(visible=False,add_book=False)
        self.wb = self.app.books.open(file)
        self.sheetNameOrIndex = sheetNameOrIndex
        self.sheet = self.wb.sheets[self.sheetNameOrIndex]
        self.keyName = keyName
        self.dataDict = excelToDict(self.wb, self.sheetNameOrIndex, keyName)
        self.startCell = self.sheet.range(startCell)
        self.formulas = getFormulaFromExcel(self.wb, sheetNameOrIndex, (self.startCell.row +1, self.startCell.column), (self.startCell.row + 1, len(self.getEmptyColumn().keys())))
        
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
        
    
    def __splitTwoDict(self, data, splitKeys):
        '''
        将一个Dict数据分割成两部分，在splitKeys中的，放入到splited中，其他的留在orig
        返回结果为：orig, spilted
        '''
        orig = data.copy()
        spilted = { k:orig.pop(k) for k in splitKeys if orig.get(k) }
        return orig, spilted

    def merge(self, *newFile, sheetNameOrIndex = 0, startCell = 'A1', forceOverWriteCols = None):
        '''
        将新制定的文件，将源文件有的且新文件也有的字段更新到源文件中；如果新文件中的记录在源文件中不存在，追加进去
        '''
        for file in newFile:
            wb = self.app.books.open(file)
            newDataDict = excelToDict(wb, sheetNameOrIndex, self.keyName, startCell)
            wb.close()
            columns = self.getEmptyColumn()
            for k, newData in newDataDict.items():
                origData = self.dataDict.get(k)
        
                if origData is None:
                    #如果文件中没有该条记录，需要将记录合并到文件，内容直接覆盖
                    self.dataDict[k] = dictOps.merge(columns.copy(), newData, UpdateOps.override)
                elif forceOverWriteCols:
                    #如果指定某些列强制覆盖，则将数据拆为两部分，不强制覆盖的继续使用为空时更新，强制覆盖的列直接覆盖
                    notForceOverwrite, forceOverwrite = self.__splitTwoDict(newData, forceOverWriteCols)
                    self.dataDict[k] = dictOps.merge(origData, notForceOverwrite, UpdateOps.whileEmpty)
                    #dataDict[k]的数据已经更新，重新合并前需要从新取数据
                    self.dataDict[k] = dictOps.merge(self.dataDict.get(k), forceOverwrite, UpdateOps.override)
                else:
                    self.dataDict[k] = dictOps.merge(origData, newData, UpdateOps.override)
            
