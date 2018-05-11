#encoding = utf - 8
#用于实现混合驱动

from . import *
from . import DataFrameWork
from testScripts.WriteTextResult import writeTextResult
from util.Log import *
from action.PageAction import *


def mixDriverRun():
    try:
        # 根据excel中sheet名称获取sheet对象
        caseIntroSheet = excelObj.getSheetByName(u"数据入口")
        # 获取“数据入口”sheet中“是否执行”列对象
        isExecuteColumn = excelObj.getColumn(caseIntroSheet,CaseIntro_IsExecute)

        funcNums = excelObj.getRowsNumber(caseIntroSheet)
        # 首次进入功能模块sheet后删除所有先前执行记录
        for myRowInFuncSheet in range(2,funcNums + 1):
            writeTextResult(caseIntroSheet, rowNo=myRowInFuncSheet,
                            colsNo="CaseIntro", testResult="")

        # 记录执行成功的功能模块数量
        successfulModuleNum = 0
        # 记录需要执行的功能模块数量
        requiredModuleNum = 0

        for Looptime, ExecuteMsg in enumerate(isExecuteColumn[1:]):

            # 获取该模块名称
            funcName = excelObj.getCellOfValue(caseIntroSheet,rowNo = Looptime + 2,colsNo = CaseIntro_funcname)

            '''
            循环遍历“数据入口”sheet中各功能模块，
            执行被设置为待执行的功能模块（即“是否待执行”为“是”的功能模块）
            '''
            if ExecuteMsg.value == "是":
                # 待执行数+1
                requiredModuleNum += 1
                # 获取“数据入口”sheet中第 Looptime+2 行模块的框架类型
                useFrameWorkName = excelObj.getCellOfValue(caseIntroSheet,
                                                           rowNo = Looptime + 2,
                                                           colsNo = CaseIntro_frameworkname)
                # 获取“数据入口”sheet中第 Looptime+2 行模块的功能&步骤sheet名
                stepSheetName = excelObj.getCellOfValue(caseIntroSheet,
                                                        rowNo = Looptime + 2,
                                                        colsNo = CaseIntro_funcsheet)
                # 获取“数据入口”sheet中第 Looptime+3 行模块的执行状态
                nextToRun = excelObj.getCellOfValue(caseIntroSheet,
                                                        rowNo = Looptime + 3,
                                                        colsNo = CaseIntro_isexecute)

                logging.info(u"********** 执行功能模块：'%s' **********" %funcName)

                '''
                判断进行关键字驱动还是数据驱动
                '''
                if useFrameWorkName == u"数据":
                    logging.info(u">> 调用数据驱动...")

                    # 获取“数据入口”sheet中第 Looptime+2 行模块的数据sheet名
                    dataSheetName = excelObj.getCellOfValue(caseIntroSheet,
                                                            rowNo = Looptime + 2,
                                                            colsNo = CaseIntro_datasheet)
                    '''
                    获取该模块的功能&步骤sheet及数据sheet对象
                    '''
                    stepSheetObj = excelObj.getSheetByName(stepSheetName)
                    dataSheetObj = excelObj.getSheetByName(dataSheetName)

                    isLastModule = False
                    if nextToRun != "是":
                        isLastModule = True
                    # 启用数据驱动框架
                    result = DataFrameWork.dataDriverRun(dataSheetObj,stepSheetObj,stepSheetName,isLastModule)

                    if result == "模块执行成功":
                        logging.info(u"功能 '%s' 执行成功" %funcName)

                        successfulModuleNum += 1
                        writeTextResult(caseIntroSheet,
                                        rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "成功")

                    elif result == "该模块不执行":
                        logging.info(u"功能 '%s' 跳过成功" %funcName)

                        successfulModuleNum += 1
                        writeTextResult(caseIntroSheet,
                                        rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "跳过")
                    else:
                        logging.info(u"功能 '%s' 执行失败" %funcName)
                        writeTextResult(caseIntroSheet,
                                        rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "失败")

                        logging.info(u"程序异常,请检查代码是否编写正确: \n%s" %result)

                elif useFrameWorkName == u"关键字":
                    logging.info(u">> 调用关键字驱动...")

                    # 获取该模块的功能 & 步骤sheet对象
                    stepSheetObj = excelObj.getSheetByName(stepSheetName)
                    # 获取该功能模块对应的功能&步骤sheet中行数
                    stepNums = excelObj.getRowsNumber(stepSheetObj)

                    # 首次进入步骤&功能sheet后删除所有先前记录
                    for myRowInStepSheet in range(2,stepNums + 1):
                        writeTextResult(stepSheetObj, rowNo=myRowInStepSheet,
                                        colsNo="CaseStep", testResult="")

                    logging.info(u"该模块共 '%s' 步" %(stepNums - 1))

                    # 定义操作步骤成功数
                    successfulStepNums = 0

                    for myRowInStepSheet in range(2,stepNums + 1):
                        # myRowInStepSheet 为功能&步骤sheet中当前步骤数据的行数
                        stepRow = excelObj.getRow(stepSheetObj,myRowInStepSheet)
                        # 获取参数
                        keyWord = stepRow[CaseStep_keyname - 1].value
                        locationType = stepRow[CaseStep_locationtype - 1].value
                        locatorExpression = stepRow[CaseStep_locatorexpression - 1].value
                        operateValue = stepRow[CaseStep_operatevalue - 1].value

                        # 用于判断是否跳出excel步骤
                        isToBreak = False

                        if isinstance(operateValue,int):
                            operateValue = str(operateValue)

                        tmpStr = "'%s','%s'" %(locationType.lower(),
                                               locatorExpression.replace("'",'"')
                                               ) if locationType and locatorExpression else ""

                        if tmpStr:
                            tmpStr += ",u'" + operateValue + "'" if operateValue else ""
                        else:
                            tmpStr += "u'" + operateValue + "'" if operateValue else ""

                        runStr = keyWord + "(" + tmpStr + ")"
                        print("********** 拼接后表达式为： ", runStr, " **********")

                        try:
                            # if operateValue != "不填":
                            eval(runStr)
                        except Exception as e:
                            print(u"********** 步骤 '%s' 执行异常 **********" %stepRow[CaseStep_stepdescribe - 1].value)
                            # 获取详细异常堆栈信息
                            errorInfo = traceback.format_exc()
                            logging.debug(u"步骤 '%s' 执行异常： \n"
                                          %stepRow[CaseStep_stepdescribe - 1].value,
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
                            successfulStepNums += 1

                            logging.info(u"步骤 '%s' 执行结束"
                                         %stepRow[CaseStep_stepdescribe - 1].value)
                            print(u"********** 执行步骤 '%s' 结束 **********" %stepRow[CaseStep_stepdescribe - 1].value)

                            # 正常结束后截图（纯关键字驱动），myValue为”是“则截图
                            myValue = stepRow[CaseStep_lockpic - 1].value

                            if myValue == '是':
                                capturePic = capture_screen()
                                writeTextResult(stepSheetObj,rowNo = myRowInStepSheet,
                                                    colsNo = "CaseStep",testResult = "成功",
                                                    CaseInfo = myValue,picPath = capturePic)
                            else:
                                writeTextResult(stepSheetObj,rowNo = myRowInStepSheet,
                                                    colsNo = "CaseStep",testResult = "成功")

                    if isToBreak:
                        logging.info(u">> 功能模块 '%s' 执行失败" %funcName)
                        writeTextResult(caseIntroSheet,rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "失败")
                        break

                    # 判断”数据入口“中功能模块执行情况
                    if successfulStepNums == stepNums - 1:
                        successfulModuleNum += 1
                        logging.info(u">> 功能模块 '%s' 执行通过" %funcName)
                        writeTextResult(caseIntroSheet,rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "成功")
                    else:
                        logging.info(u">> 功能模块 '%s' 执行失败" %funcName)
                        writeTextResult(caseIntroSheet,rowNo = Looptime + 2,
                                        colsNo = "CaseIntro",testResult = "失败")



            else:
                # 将“是否待执行”记录不为“是”的行执行结果清空.无效果
                writeTextResult(caseIntroSheet,rowNo = Looptime + 2,colsNo = "CaseIntro",testResult = "")

                logging.info(u"********** 功能模块 '%s' 被设置为忽略执行 **********" %funcName)


        # if myRow:       # 在功能模块执行结束并且成功再在数据表写整行数据的使用状态
        #     writeTextResult(sheetObj=dataSheetObj, rowNo=myRow,
        #                     colsNo="DataSource", testResult="成功",
        #                     dataUse="已使用")

        logging.info(u">> 共 %d 个功能模块, %d 个需要被执行,成功执行 %d 个 \n\n"

                     %(len(isExecuteColumn) - 1,requiredModuleNum,successfulModuleNum))
        # close_browser()

    except Exception as e:
        logging.info(u">> 共 %d 个功能模块, %d 个需要被执行,成功执行 %d 个 \n"
                     %(len(isExecuteColumn) - 1,requiredModuleNum,successfulModuleNum))
        logging.debug(u"程序异常,请检查代码是否编写正确: \n%s" %traceback.format_exc() + "\n")
        # close_browser()
