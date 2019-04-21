'''
Created on 2019年4月13日

@author: Administrator
'''

from .updateOps import UpdateOps

def update(orig, newData, updateOps = UpdateOps.whileEmpty):
    x = orig.copy()
    for key in x.keys():
        #只有源数据为空时才覆盖
        if updateOps == UpdateOps.whileEmpty:
            origValue = x.get(key)
            if (origValue is None or str(origValue).strip() == ''):
                x[key] = newData.get(key)
        #覆盖时，要先判断新数据是否有key，如果没有，则不需要覆盖
        elif updateOps == UpdateOps.override and key in newData:
            x[key] = newData.get(key) 
        else:
            pass   
    
    return x

def appendNotExistKeys(orig, newData):
    x = { k:v for k,v in newData.items() if k not in orig}
    return { **orig, **x}

def merge(orig, toBeComb, updateOps = UpdateOps.whileEmpty, joinFlag = False):
    #合并时覆盖且需要连接，直接使用dict update方法
    if updateOps == UpdateOps.override and joinFlag:
        return { **orig, **toBeComb}
    
    x = update(orig, toBeComb, updateOps)
    if joinFlag:
        #以 | 保证orig的优先顺序，保证updateOps可信
        x =appendNotExistKeys(x, toBeComb)
    
    return x

if __name__ == '__main__':
    #输出结果分别为
#===============================================================================
#     {'a': 1, 'b': 2}
#     {'a': 1, 'b': 2, 'c': 4}
#     {'a': 1, 'b': 3}
#     {'a': 1, 'b': 3, 'c': 4}
#===============================================================================
    
    x = {'a': 1, 'b': 2}
    y = {'b': 3, 'c': 4}
    
    z = merge(x, y, updateOps = UpdateOps.whileEmpty, joinFlag = False)
    print(z)
    z = merge(x, y, updateOps = UpdateOps.whileEmpty, joinFlag = True)
    print(z)
    z = merge(x, y, updateOps = UpdateOps.override, joinFlag = False)
    print(z)
    z = merge(x, y, updateOps = UpdateOps.override, joinFlag = True)
    print(z)
    
    #===========================================================================
    # {'a': 1, 'b': 3, 'c': 4}
    # {'a': 1, 'b': 2, 'c': 'c'}
    # {'a': 1, 'b': 2, 'c': 'c', 4: 4}
    # {'a': 1, 'b': 2, 'c': 'd'}
    # {'a': 1, 'b': 2, 'c': 'd', 4: 4}
    #===========================================================================
    k = {'a': None, 'b': '','c':'c'}
    l = {'a': 1, 'b': 2, 'c':'d', 4:4}
    z = merge(k, l, updateOps = UpdateOps.whileEmpty, joinFlag = False)
    print(z)
    z = merge(k, l, updateOps = UpdateOps.whileEmpty, joinFlag = True)
    print(z)
    z = merge(k, l, updateOps = UpdateOps.override, joinFlag = False)
    print(z)
    z = merge(k, l, updateOps = UpdateOps.override, joinFlag = True)
    print(z)
    
    pass