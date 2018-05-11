#encoding = utf - 8
# 用于定义框架中用到的全局变量

import os

'''
配置全局变量后无需设置driver路径
'''
# ieDriverFilePath = "C:\Program Files\Internet Explorer\IEDriverServer.exe"
# chromeDriverFilePath = "C:\Python34\chromedriver.exe"
# firefoxDriverFilePath = "C:\Python34\geckodriver.exe"

# 获取当前文件所在目录的绝对路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 异常图片存放目录
screenPicturesDir = parentDirPath + u"\\processpictures\\"

# # 获取存放页面元素定位表达式文件的绝对路径,纯数据驱动中才需要，较繁琐
# parentElementLocatorPath = parentDirPath + u"\\config\\PageElementLocator.ini"

# 获取数据文件存放的绝对路径,数据文件后期更换，TODO
dataFilePath = parentDirPath + u"\\testData\\数据汇总.xlsx"

'''
数据文件excel中，每列对应的数字编号，
data_isExecute用于选中当前轮次执行数据，标记列号为字母
其余数据列标记为数字
'''
# 数据入口
CaseIntro_IsExecute = 'H'

CaseIntro_funcname = 3
CaseIntro_frameworkname = 5
CaseIntro_funcsheet = 6
CaseIntro_datasheet = 7
CaseIntro_isexecute = 8
CaseIntro_runtime = 9
CaseIntro_runresult = 10

# 功能&步骤模块
CaseStep_stepdescribe = 2
CaseStep_keyname = 3
CaseStep_locationtype = 4
CaseStep_locatorexpression = 5
CaseStep_operatevalue = 6
CaseStep_isreturned = 7
CaseStep_runtime = 8
CaseStep_runresult = 9
CaseStep_errorinfo = 10
CaseStep_lockpic = 11
CaseStep_errorpic = 12

# 数据模块，视情况须有所补充，TODO
DataSource_IsExecute = 'C'

DataSource_isexecute = 3
DataSource_runtime = 4
DataSource_runresult = 5
DataSource_processdata = 12
DataSource_finaldata = 13
