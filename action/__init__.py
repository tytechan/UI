#encoding = utf - 8

from action.PageAction import *
from util.ParseExcel import ParseExcel
from config.VarConfig import *

# 创建解析excel对象
excelObj = ParseExcel()
# 将excel数据文件加载至内存
excelObj.loadWorkBook(dataFilePath)