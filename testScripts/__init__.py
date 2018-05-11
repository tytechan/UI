#encoding = utf - 8

from action.PageAction import *
from util.ParseExcel import ParseExcel
from config.VarConfig import *
import time
import traceback

# # 设置编码环境为utf-8
# import sys
# reload(sys)
# sys.setdefaultencoding("utf-8")

# 创建解析excel对象
excelObj = ParseExcel()
# 将excel数据文件加载至内存
excelObj.loadWorkBook(dataFilePath)