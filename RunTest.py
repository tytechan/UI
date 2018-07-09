#encoding = utf - 8
# 用于驱动执行

from testScripts.MixFrameWork import mixDriverRun
from action.PageAction import *

if __name__ == "__main__":
    # 可添加循环并控制循环次数，以在同一流程测试多组数据，TODO
    looptime = 1
    for i in range(1,looptime+1):
        print("********** 第",i,"次执行 **********")
        mixDriverRun()
        # close_browser()