'''
Created on 2019年4月13日

@author: Administrator
'''

from enum import Enum, unique


@unique
class UpdateOps(Enum):
    whileEmpty = 1 #值为空时才允许更新
    override = 2 #直接用新值覆盖

