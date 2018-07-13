#encoding = utf - 8
# 用于向excel中写入执行结果

from . import *

# 执行完成后，向excel中写入执行信息
def writeTextResult(sheetObj,rowNo,colsNo,testResult,CaseInfo = None,picPath = None,dataUse = None,returnValue = None):
    # 控制执行结果颜色控制字典
    colorDict = {"成功":"green",
                 "失败":"red",
                 "跳过":"green",
                 "":None}

    # 以下字典用于区分写入sheet类型
    colsDict = {
        "CaseIntro":[CaseIntro_runtime,CaseIntro_runresult],
        "CaseStep":[CaseStep_runtime,CaseStep_runresult],
        "DataSource":[DataSource_runtime,DataSource_runresult]
    }

    try:
        # 先初始化表格颜色，否则颜色变换会出错
        excelObj.writeCell(sheetObj,content = testResult,rowNo = rowNo,colsNo = colsDict[colsNo][1],style = colorDict[""])
        # 写入执行结果
        excelObj.writeCell(sheetObj,content = testResult,rowNo = rowNo,colsNo = colsDict[colsNo][1],style = colorDict[testResult])

        if testResult == "":
            # 清空时间单元格
            excelObj.writeCell(sheetObj,content = "",rowNo = rowNo,colsNo = colsDict[colsNo][0])
        else:
            # 写入执行时间
            excelObj.writeCellCurrentTime(sheetObj,rowNo = rowNo,colsNo = colsDict[colsNo][0])

        if colsNo == "CaseStep":
            if CaseInfo and picPath:
                # 在功能 & 步骤表格中写入截图文件路径
                excelObj.writeCell(sheetObj,content = picPath,rowNo = rowNo,colsNo = CaseStep_errorpic)
                print("**********锁定截图**********")

                if CaseInfo != "是":
                    # 在功能 & 步骤表格中写入异常信息
                    excelObj.writeCell(sheetObj,content = CaseInfo,rowNo = rowNo,colsNo = CaseStep_errorinfo)
            else:
                # 在功能 & 步骤表格中清空异常信息单元格
                excelObj.writeCell(sheetObj,content = "",rowNo = rowNo,colsNo = CaseStep_errorinfo)
                excelObj.writeCell(sheetObj,content = "",rowNo = rowNo,colsNo = CaseStep_errorpic)

        # 返回值格可以填数据表列字母，根据列字母，将返回值存于数据表，rowNo行的指定列
        # 若需返回多个值，则需返回字符串，每个值以"[]"为分隔，返回值格填写对应的多个位置信息，也以"[]"分隔
        if colsNo == "DataSource":
            excelObj.writeCell(sheetObj, content=dataUse, rowNo=rowNo, colsNo=DataSource_isexecute)
            # 返回值格式：{'位置信息':'返回值', '位置信息':'返回值'}
            for (k, v) in returnValue.items():
                # 位置信息
                position = k
                # 返回值内容
                return_value = v
                # 将返回值存于“过程池”或“结果池”
                if position == "过程":
                    excelObj.writeCell(sheetObj,content=return_value,rowNo=rowNo,colsNo=DataSource_processdata)
                elif position == "结果":
                    excelObj.writeCell(sheetObj,content=return_value,rowNo=rowNo,colsNo=DataSource_finaldata)

                # 将返回值存于数据表指定列，TODO .04:通过列号修改为通过列表头返回值
                # elif position.encode('utf-8').isalpha():
                #     position_coordinate = "%s%d" %(position, rowNo)
                #     excelObj.writeCell(sheetObj, content=return_value, coordinate=position_coordinate)

                elif position.startswith("#") == True:
                    myColumn = 1
                    for myBox in excelObj.getRow(dataSourceSheetObj, 1):
                        if myBox.value == None:
                            break
                        elif myBox.value == position.replace("#",""):
                            excelObj.writeCell(sheetObj, content=return_value,rowNo=rowNo,colsNo=myColumn)
                            break
                        myColumn += 1


    except Exception as e:
        print(u"********** excel写入执行结果失败，错误信息为： **********")
        print(traceback.print_exc())