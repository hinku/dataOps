def __splitTwoDict(data, splitKeys):
        '''
        将一个Dict数据分割成两部分，在splitKeys中的，放入到splited中，其他的留在orig
        返回结果为：orig, spilted
        '''
        orig = data.copy()
        spilted = { k:copyData.pop(k) for k in splitKeys if copyData.get(k) }
        return orig, spilted


a = {1:1, 2:2, 3:3}

o,s = __splitTwoDict(a, [2])

print(o)
print(s)