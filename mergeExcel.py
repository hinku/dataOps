'''
Created on 2019年4月13日

@author: Administrator
'''

import os
import sys
import json
from dataOps.excelOps import ExcelOps
from logs.logger import logger

def toAbspath(file):
    """
        将文件转成绝对路径，如果已经是绝对路径，则直接返回，否则返回当前所在的绝对路径
    """
    if file == os.path.abspath(file):
        return file
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), file)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        cfgFile = sys.argv[0]
    else:
        cfgFile = 'merge.json'

    path = toAbspath(cfgFile)
    if not os.path.isfile(path):
        logger.error('%s is not a file' % path)
        sys.exit(1)
    
    with open(path, 'r', encoding = 'utf-8') as f:
        cfg = json.load(f)
        
    logger.info('merge excel')
    xls = ExcelOps(toAbspath(cfg['dstFile']), sheetNameOrIndex = cfg.get('dtsFileSheet', 0), keyName = cfg['keyName'], startCell = cfg.get('dtsFileStartCell', 'A1'))
    print(xls.dataDict)
    xls.merge(toAbspath(cfg['srcFile']),  sheetNameOrIndex = cfg.get('srcFileSheet', 0), startCell = cfg.get('srcFileStartCell', 'A1'), forceOverWriteCols = cfg.get('forceOverWriteCols'))
    print(xls.dataDict)
    xls.flush()
        