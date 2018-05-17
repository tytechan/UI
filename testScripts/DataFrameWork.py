#encoding = utf - 8
# 用于实现数据框架驱动

from . import *
from testScripts.WriteTextResult import *
from util.Log import *

def  dataDriverRun(dataSourceSheetObj,stepSheetObj,stepSheetName,isLastModule,funcName):
    '''
    :param dataSourceSheetObj: 数据模块sheet对象
    :param stepSheetObj: 功能&步骤模块sheet对象
    :return:
    '''
    try:
        # 获取数据模块sheet中“数据状态”列对象
        dataIsExecuteColumn = excelObj.getColumn(dataSourceSheetObj,DataSource_IsExecute)
        # 获取功能&步骤模块sheet中存在数据区域的行数
        stepRowNums = excelObj.getRowsNumber(stepSheetObj)
        # print("功能&步骤sheet中行数为：",stepRowNums)

        # 每次进入步骤&功能sheet后删除所有先前记录
        for myRowInStepSheet in range(2, stepRowNums + 1):
            writeTextResult(stepSheetObj, rowNo=myRowInStepSheet,
                            colsNo="CaseStep", testResult="")
        '''
        requiredDataNo和successStepNo用于判断执行结束标志
        '''
        # 记录成功执行的数据行数
        successDataNo = 0
        # 记录待执行的数据行数
        requiredDataNo = 0

        for Looptime, ExcuteMsg in enumerate(dataIsExecuteColumn[1:]):
            # print("***** Looptime:",Looptime," ***** ExcuteMsg:",ExcuteMsg,"： ",ExcuteMsg.value)

            # 用于判断是否跳出excel步骤
            isToBreak = False
            # 先在数据模块sheet中遍历，判断该行数据是否已执行
            if ExcuteMsg.value != "已使用":
                print("********** 开始调用第 ",Looptime + 2," 行数据 **********")

                # 该功能模块不执行时，跳出循环，TODO
                # 在数据sheet中新增列，判断功能模块是否跳过
                jumpToBreak = False
                # print(excelObj.getRow(dataSourceSheetObj,1))
                myColumnNum = 1
                for myBox in excelObj.getRow(dataSourceSheetObj,1):
                    # 获取数据sheet中与功能sheet同名的列，并判断（Looptime + 2，myColumnNum）的值是否为“跳过”
                    if myBox.value == funcName:
                        jumpValue = excelObj.getCellOfValue(dataSourceSheetObj,rowNo=Looptime+2,colsNo=myColumnNum)
                        print("********** 第",Looptime+2,"行",myColumnNum,"列的判断跳出标志位为：",
                              jumpValue," **********")
                        if jumpValue == stepSheetName:
                            pass
                        else:
                            jumpToBreak = True
                        break
                    myColumnNum += 1
                if jumpToBreak:
                    # 若对应表格内值不为“执行”，则不用执行该模块
                    logging.info(u">> 跳过该模块所有步骤...")
                    break


                logging.info(u">> 开始调用数据表...")

                requiredDataNo += 1
                # 定义执行成功步骤数变量
                successStepNo = 0

                # 遍历功能&步骤sheet中存在数据的所有行
                for myRowInStepSheet in range(2,stepRowNums + 1):
                    # 获取sheet中第 myRowInStepSheet 行对象
                    rowObj = excelObj.getRow(stepSheetObj,myRowInStepSheet)
                    # 获取关键字作为调用的函数名
                    keyWord = rowObj[CaseStep_keyname - 1].value
                    # 获取定位方式
                    locationType = rowObj[CaseStep_locationtype - 1].value
                    # 获取定位表达式
                    locatorExpression = rowObj[CaseStep_locatorexpression - 1].value
                    # 获取操作值
                    operateValue = rowObj[CaseStep_operatevalue - 1].value
                    # 获取判断是否有返回值标志位
                    isReturnedValue = rowObj[CaseStep_isreturned - 1].value

                    if operateValue:
                        if isinstance(operateValue,int):
                            print("*********** 直接通过关键字驱动输入的值为： ",operateValue," ***********")
                            operateValue = str(operateValue)
                            print("数值型operateValue值为：",operateValue)
                        if operateValue and operateValue.encode('utf-8').isalpha():
                            '''
                            operateValue不为空，且所有字符均为字母，则说明为调用情况
                            '''
                            print("字母型operateValue值为：",operateValue)
                            coordinate = operateValue + str(Looptime + 2)
                            print("获取数据坐标coordinate为：",coordinate)
                            operateValue = excelObj.getCellOfValue(dataSourceSheetObj,coordinate = coordinate)
                            operateValue = str(operateValue)

                            print("********** 数据表sheet中调用单元格为： ",coordinate," 对应值为： ",operateValue," **********")

                        if operateValue.startswith("&"):
                            '''
                            若以“*”开头，则说明该字符串为纯英文，但不是调用数据sheet情况
                            '''
                            operateValue = str(operateValue.split("&")[1])

                    # 拼接字符串获得需要执行的python表达式，以对应 PageAction.py 中对应函数方法
                    tmpStr = "'%s','%s'" %(locationType.lower(),
                                           locatorExpression.replace("'",'"')
                                           ) if locationType and locatorExpression else ""
                    if tmpStr:
                        tmpStr += ",u'" + operateValue + "'" if operateValue else ""
                    else:
                        tmpStr += "u'" + operateValue + "'" if operateValue else ""

                    runStr = keyWord + "(" + tmpStr + ")"

                    print("********** 拼接后表达式为： ",runStr," **********")

                    try:
                        if operateValue != "不填":
                            # 执行表达式，并返回结果
                            if isReturnedValue:     # 该步骤有返回值
                                valueReturned = eval(runStr)
                                print("********** 返回值为：",valueReturned," **********")
                            else:       # 该步骤无返回值
                                eval(runStr)
                    except Exception as e:
                        print(u"********** 步骤 '%s' 执行异常 **********" %rowObj[CaseStep_stepdescribe - 1].value)

                        errorInfo = traceback.format_exc()
                        logging.info(u"步骤 '%s' 执行异常： \n"
                                      %rowObj[CaseStep_stepdescribe - 1].value,
                                      errorInfo)
                        # 截取异常截图
                        capturePic = capture_screen()
                        writeTextResult(stepSheetObj,rowNo = myRowInStepSheet,
                                        colsNo = "CaseStep",testResult = "失败",
                                        CaseInfo = str(errorInfo),picPath = capturePic)

                        # 该步骤失败后，模块运行终止
                        isToBreak = True
                        break
                    else:
                        successStepNo += 1
                        logging.info(u"步骤 '%s' 执行结束"
                                     %rowObj[CaseStep_stepdescribe - 1].value)
                        print(u"********** 执行步骤 '%s' 结束 **********" %rowObj[CaseStep_stepdescribe - 1].value)

                        # 正常结束后截图（混合驱动），myValue为”是“则截图，TODO
                        myValue = rowObj[CaseStep_lockpic - 1].value
                        if myValue == '是':
                            capturePic = capture_screen()
                            writeTextResult(stepSheetObj,rowNo = myRowInStepSheet,
                                                colsNo = "CaseStep",testResult = "成功",
                                                CaseInfo = str(myValue),picPath = capturePic)
                        else:
                            writeTextResult(stepSheetObj, rowNo=myRowInStepSheet,
                                            colsNo="CaseStep", testResult="成功")

                        # 判断返回值情况
                        if isReturnedValue:
                            valueReturned = isReturnedValue + valueReturned
                            writeTextResult(dataSourceSheetObj,rowNo = Looptime + 2,
                                            colsNo = "DataSource",testResult = "失败",
                                            returnValue = valueReturned)
                            valueReturned = None

                if isToBreak:       #执行失败跳出标志
                    writeTextResult(sheetObj = dataSourceSheetObj,rowNo = Looptime + 2,
                                    colsNo = "DataSource",testResult = "失败",
                                    dataUse = "未使用")
                    break

                if stepRowNums == successStepNo + 1:
                    '''
                    如果成功执行步骤数 successStepNo 等于表中给出的步骤数，
                    说明第 Looptime+2 行数据执行通过，则写入成功信息                     
                    '''
                    if isLastModule:
                        writeTextResult(sheetObj = dataSourceSheetObj,rowNo = Looptime + 2,
                                        colsNo = "DataSource",testResult = "成功",
                                        dataUse = "已使用")
                    successDataNo += 1
                    # 若该行数据执行成功，则不再执行下一行数据，待确定可否控制是否循环执行所有，TODO！！！
                    break
                else:
                    # 写入失败信息
                    writeTextResult(sheetObj = dataSourceSheetObj,rowNo = Looptime + 2,
                                    colsNo = "DataSource",testResult = "失败",
                                    dataUse="未使用")


                # # 用于控制数据sheet每执行一条“未使用”数据则跳出循环继续下一个步骤，须格外注意，TODO
                # # 通过循环清理sheet中之前的执行记录时，不可break
                # break

            # else:
            #     # 将“数据状态”记录不为“未使用”的行执行结果清空
                # writeTextResult(sheetObj = dataSourceSheetObj,rowNo = Looptime + 2,colsNo = "DataSource",testResult = "")

        if requiredDataNo == successDataNo:
            '''
            成功执行数=待执行数 时，表示执行成功或跳出模块
            '''
            if jumpToBreak:
                return "该模块不执行"
            else:
                return "模块执行成功"
        else:
            # 表示数据驱动失败
            return errorInfo

    except Exception as e:
        raise e