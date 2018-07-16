#encoding = utf - 8
# 用于获取当前日期及时间，并创建异常截图存放目录

import time,os
from datetime import datetime
from config.VarConfig import screenPicturesDir

def getCurrentDate():       #获取当前日期
    timeTup = time.localtime()
    currentDate = str(timeTup.tm_year) + "-" + \
        str(timeTup.tm_mon) + "-" + str(timeTup.tm_mday)
    return currentDate

def getCurrentTime():       #获取当前时间，确保日期中无非法字符，否则无法生成截图文件
    timeStr = datetime.now()
    nowTime = timeStr.strftime('%H-%M-%S.%f')
    return nowTime

def createCurrentDateDir():
    dirName = os.path.join(screenPicturesDir,getCurrentDate())
    if not os.path.exists(dirName):
        os.makedirs(dirName)
    return dirName


if __name__ == "__main__":
    print('当前日期为',getCurrentDate())
    print('当前时间为',getCurrentTime())
    print('截图存放目录为',createCurrentDateDir())