#encoding = utf - 8
# 用于驱动执行


import os,configparser as cparser
from config.global_config import *



def createProcessPicDir():
    '''返回本流程的截图存放路径'''
    from util.DirAndTime import getCurrentDate
    from util.DirAndTime import getCurrentTime
    from config.VarConfig import screenPicturesDir

    myDirName = os.path.join(screenPicturesDir,getCurrentDate())
    if not os.path.exists(myDirName):
        os.makedirs(myDirName)
    dirName = myDirName + "\\" + getCurrentTime()
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName


def createScopeOfExecution():
    '''生成执行范围json数据'''
    base_dir = str(os.path.dirname(__file__)).replace('\\', '/')
    file_path = base_dir + "/testData/Scope_of_execution.ini"

    cf = cparser.ConfigParser()
    cf.read(file_path, encoding='UTF-8')

    executeJson = \
        {
            "执行范围": []
        }

    for section in cf.sections():
        numToExecute = cf.get(section,"numToExecute")
        filePath = base_dir + cf.get(section,"filePath")

        executeJsonToAdd = \
            {
                "执行流程数": numToExecute,
                "文件路径": filePath
            }

        if int(executeJsonToAdd.get("执行流程数")) > 0:
            executeJson.get("执行范围").append(executeJsonToAdd)
    # print(executeJson)
    return executeJson


def mainRunningPart():
    # i为同一流程当前的执行次数
    from testScripts.MixFrameWork import mixDriverRun

    picDir = createProcessPicDir()
    mixDriverRun(picDir)


def toExit():
    from action.PageAction import close_browser, closeSAP, closeAllSession
    # closeSAP()                # 关闭所有打开的sap进程
    closeAllSession()         # 关闭本流程的所有sap进程
    # close_browser()           # 关闭本流程浏览器进程
    # pass


if __name__ == "__main__":
    # 根据配置文件“/testData/Scope_of_execution.ini”，生成执行范围json：executeJson
    executeJson = createScopeOfExecution()

    myZip = []
    for j in range(0,len(executeJson.get("执行范围"))):
        create_dict()
        numToExecute = int(executeJson["执行范围"][j]["执行流程数"])
        filePath = executeJson["执行范围"][j]["文件路径"]
        myZip.append(numToExecute)


        set_value("FILEPATH",None)
        set_value("FILEPATH",filePath)

        for i in range(1, numToExecute + 1):
            if i == 1:
                set_value("已用数据行", 0)
            set_value("本流程执行次数", i)
            print("********** 第",i,"次执行 “%s” 文件中流程 **********" %(filePath.split("/")[-1]))
            mainRunningPart()
            set_value("已用数据行",get_value("可用数据行"))
            toExit()

    mySum = sum(myZip)
    print("********** 共执行自动化流程 %d 条 **********" %mySum)