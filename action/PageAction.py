#encoding = utf - 8
# 用于存放页面操作方法

from selenium import webdriver
from util.ObjectMap import *
from util.ClipboardUtil import Clipboard
from util.KeyBoardUtil import KeyboardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import win32con
import win32gui


from util.ParseExcel import ParseExcel
from config.VarConfig import *
from openpyxl.writer.excel import ExcelWriter
from . import *


from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options

# 定义全局driver变量
driver = None
# 全局的等待类实例对象
waitUtil = None

def open_browser(browserName,*arg):        #打开浏览器
    global driver,waitUtil
    try:
        if browserName.lower() == 'ie':
            driver = webdriver.Ie()
        elif browserName.lower() == 'chrome':
            # # 创建chrome浏览器的一个options实例对象
            # chrome_options = Options()
            # # 添加屏蔽--ignore--certificate-errors提示信息的设置参数项
            # chrome_options.add_experimental_option(
            #     "excludeSwitches",
            #     ["ignore-certificate-errors"]
            # )
            # driver = webdriver.Chrome(executable_path = chromeDriverFilePath,chrome_options = chrome_options)

            driver = webdriver.Chrome()
        else:
            driver = webdriver.Firefox()
        # driver对象创建成功，创建等待类实例对象
        waitUtil = WaitUtil(driver)
    except Exception as e:
        raise e

def visit_url(url,*arg):        #访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception as e:
        raise e

def close_browser(*arg):        #关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception as e:
        raise e

def sleep(sleepSeconds,*arg):       #强制等待
    try:
        time.sleep(int(sleepSeconds))
    except Exception as e:
        raise e

def clear(locationType,locatorExpression,*arg):     #清除输入框默认内容
    global driver
    try:
        findEleByDetail(driver,locationType,locatorExpression).clear()
    except Exception as e:
        raise e

def sendkeys_To_Obj(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    try:
        element = findEleByDetail(driver,locationType,locatorExpression)
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        raise e

def click_Obj(locationType, locatorExpression, *arg):       #点击页面元素
    global driver
    try:
        findEleByDetail(driver, locationType, locatorExpression).click()
    except Exception as e:
        raise e

def assert_string_in_pagesourse(assertstring,*arg):     #断言当前页面是否存在指定字段
    global driver
    try:
        driver.implicitly_wait(10)
        assert assertstring in driver.page_source, \
            u"在当前页面未找到字段：%s" %assertstring
        # startTime = time.time()
        # for i in range(20):
        #     myTime = time.time() - startTime
        #     if assert assertstring in driver.page_source:
        #         break
        #     if myTime <= 10:
        #         sleep(0.5)

    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e

def asser_title(titleStr,*arg):     #断言判断当前页面标题是否存在指定字段
    global driver
    try:
        assert titleStr in driver.title, \
            u"当前不存在标题为 %s 的页面" % titleStr
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e

def getTitle(*arg):     #获取页面标题
    global driver
    try:
        return driver.title
    except Exception as e:
        raise e

def getPageSource(*arg):        #获取页面源码
    global driver
    try:
        return driver.page_source
    except Exception as e:
        raise e

def switch_to_frame(locationType,frameLocatorExpression,*arg):      #切换进入frame
    global driver
    try:
        driver.switch_to.frame(findEleByDetail(driver,locationType,frameLocatorExpression))
    except Exception as e:
        print('未找到指定frame')
        raise e

def switch_to_default_content(*arg):        #切出frame，回到默认对话框中
    global driver
    try:
        return driver.switch_to.default_content()
    except Exception as e:
        raise e

def paste_string(pasteString,*arg):     #模拟 ctrl+v
    try:
        Clipboard.setText(pasteString)
        time.sleep(2)
        KeyboardKeys.twoKeys("ctrl","v")
    except Exception as e:
        raise e

def press_key(mykey,*arg):        #模拟单按键，如： "tab"、"enter"
    try:
        KeyboardKeys.oneKey(mykey)
    except Exception as e:
        raise e

def maximize_browser():     #窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception as e:
        raise e

def capture_screen(*arg):       #截图
    global driver
    # 获取当前时间，精确到秒
    currentTime = getCurrentTime()
    # 拼接一场图片保存的绝对路径及名称
    picNameAndPath = str(createCurrentDateDir()) + "\\" + str(currentTime) + ".png"
    try:
        # 截屏，并保存为本地图片
        driver.get_screenshot_as_file(picNameAndPath.replace('\\',r'\\'))
        # print("picNameAndPath 为：",picNameAndPath.replace('\\',r'\\'))
    except Exception as e:
        raise e
    else:
        return picNameAndPath

def waitPresenceOfElementLocated(locationType,locatorExpression,*arg):
    '''
    显示等待页面元素出现在DOM中，但不一定可见，存在则返回页面元素对象
    :param locationType: 定位方法
    :param locatorExpression: 定位表达式
    :param arg:
    :return: 页面元素对象
    '''
    global waitUtil
    try:
        waitUtil.presenceOfElementLocated(locationType,locatorExpression)
    except Exception as e:
        raise e

def waitVisibilityOfElementLocated(locationType,locatorExpression,*arg):
    '''
    显式等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象
    :param locationType: 定位方法
    :param locationExpression: 定位表达式
    :param args:
    :return: None
    '''
    global waitUtil
    try:
        waitUtil.visibilityOfElementLocated(locationType,locatorExpression)
    except Exception as e:
        raise e

def waitFrameToBeAvailableAndSwitchToIt(locationType,locatorExpression,*arg):
    '''
    检查frame是否存在，存在则切换到frame控件中
    :param locationType: 定位方法
    :param LocationExpression: 定位表达式
    :param args:
    :return: None
    '''
    global waitUtil
    try:
        waitUtil.frameToBeAvailableAndSwitchToIt(locationType,locatorExpression)
    except Exception as e:
        raise e

def moveToElement(locationType,locatorExpression,*arg):        #鼠标移动到指定元素
    global driver
    try:
        element = findEleByDetail(driver, locationType, locatorExpression)
        MoveToEle(driver,element)
    except Exception as e:
        raise e


# 针对partial_link_text、link_text、css_selector报错Unsupported locator strategy封装单独关键字

def specObjClear(locationType,locatorExpression,*arg):     #清除输入框默认内容，暂时弃用
    global driver
    try:
        findElebyMethod(driver,locationType,locatorExpression).clear()
    except Exception as e:
        raise e

def sendkeys_To_SpecObj(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    try:
        element = findElebyMethod(driver,locationType,locatorExpression)
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        raise e

def click_SpecObj(locationType, locatorExpression, *arg):       #点击页面元素
    global driver
    try:
        findElebyMethod(driver, locationType, locatorExpression).click()
    except Exception as e:
        raise e

def moveToElement(locationType,locatorExpression,*arg):        #鼠标移动到指定元素
    global driver
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        element = findElebyMethod(driver, locationType, locatorExpression)
        ActionChains(driver).move_to_element(element).perform()
        # print(element.get_attribute("LINK_TEXT"))
        # MoveToEle(driver,element)
    except Exception as e:
        raise e

def init_Mouse(*arg):       # 初始化鼠标位置
    try:
        moveMouse(100,10)
    except Exception as e:
        raise e

def highlightElement(locationType,locatorExpression,*arg):     # 高亮元素
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        highlight(driver,element)
    except Exception as e:
        raise e

def highlightElements(locationType,locatorExpression,*arg):     # 高亮元素
    global driver
    try:
        elements = findElesbyMethod(driver, locationType, locatorExpression)
        print("********** 共有 ",len(elements)," 个高亮元素 **********")
        for i in elements:
            highlight(driver,i)
            sleep(0.5)
    except Exception as e:
        raise e

def whichIsEnabled(locationType,locatorExpression,*arg):    # 判断元素列表中哪些为可操作的
    global driver
    try:
        elements = findElesbyMethod(driver, locationType, locatorExpression)
        print("********** 共有 ",len(elements)," 个待判断可操作的元素 **********")
        time = 1
        for i in elements:
            if i.is_enabled():
                print("********** 第 ",time," 个元素为可操作 **********")
            time += 1
    except Exception as e:
        raise e

def whichIsDisplayed(locationType,locatorExpression,*arg):    # 判断元素列表中哪些为可见的
    global driver
    try:
        elements = findElesbyMethod(driver, locationType, locatorExpression)
        print("********** 共有 ",len(elements)," 个待判断可见的元素 **********")
        time = 1
        for i in elements:
            if i.is_displayed():
                print("********** 第 ",time," 个元素为可见")
            time += 1
    except Exception as e:
        raise e

def loadPage(*arg):     # 设置页面加载时间
    global driver
    try:
        sleep(1)
        driver.set_page_load_timeout(10)
        sleep(1)
    except TimeoutError as e:
        print("********** 等待页面加载超时 **********")
        raise TimeoutError(e)

def pageKeySimulate(locationType,locatorExpression,keyType,*arg):      # 模拟键盘
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        if keyType == "page_down":
            element.send_keys(Keys.PAGE_DOWN)
        if keyType == "page_home":
            element.send_keys(Keys.HOME)
        if keyType == "page_end":
            element.send_keys(Keys.END)
    except Exception as e:
        raise e

def uploadFile_x1(fileName,*arg):      # 上传文件，文件路径为testData路径（使用失败），TODO
    global driver
    try:
        myPath = parentDirPath + u"\\testData\\" + fileName
        dialog = win32gui.FindWindow('#32770', u'文件上传')  # 对话框
        ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, 'ComboBoxEx32', None)
        ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, 'ComboBox', None)
        Edit = win32gui.FindWindowEx(ComboBox, 0, 'Edit', None)  # 上面三句依次寻找对象，直到找到输入框Edit对象的句柄
        button = win32gui.FindWindowEx(dialog, 0, 'Button', None)  # 确定按钮Button

        win32gui.SendMessage(Edit, win32con.WM_SETTEXT, None, myPath)  # 往输入框输入绝对地址
        win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)  # 按button
    except Exception as e:
        raise e

def uploadFile_x2(fileName,*arg):  # 模拟键盘上传文件，文件路径为testData路径（使用失败），TODO
    global driver
    try:
        myPath = parentDirPath + u"\\testData\\" + fileName
        print("********** 上传文件的据对路径为：",myPath," **********")
        Clipboard.setText(myPath)
        # 将文件路径写入剪贴板
        Clipboard.getText()
        KeyboardKeys.twoKeys("ctrl","v")
        sleep(1)
        KeyboardKeys.oneKey("enter")
        sleep(1)
    except Exception as e:
        raise e

def runProcessFile(fileName,*arg):    # autoit上传文件
    global driver
    try:
        myPath = parentDirPath + u"\\fileHandle\\" + fileName
        print("********** 调用文件的据对路径为：",myPath," **********")
        os.system(myPath)
        sleep(1)
    except Exception as e:
        raise e

def randomNum(len,*arg):        # 生成指定长度的随机数（长度>=6）
    try:
        import random,datetime
        list_num = ['0','1','2','3','4','5','6','7','8','9']
        result = []
        len = int(len)
        mydate = datetime.datetime.now().strftime('%Y%m%d')
        result.append(mydate[2:])
        for i in range(0,len-6):
            result.append(random.choice(list_num))
        return "".join(result)
    except Exception as e:
        raise e

def writeContracNum(myInfo,*arg):
    # 该方法加断点时可往excel中写值成功，不加断点则写不进去，randomContracNum 方法暂时启用，TODO
    try:
        # ParseExcel().randomContracNum(myInfo)
        randContractNum = myInfo + randomNum(9)
        return randContractNum
    except Exception as e:
        raise e

def getAttribute(locationType,locatorExpression,attributeType,*arg):        # 获取页面元素属性值
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        attributeValue = element.get_attribute(attributeType)
        return attributeValue
    except Exception as e:
        raise e

def SelectValues(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    try:
        el = Select(findEleByDetail(driver,locationType,locatorExpression))
        el.select_by_visible_text(inputContent)
    except Exception as e:
        raise e

def setValueByTextAside(textAside,inputContent,*arg):       # 根据输入框旁边的字段定位并向输入框输值,待整理参数，TODO
    global driver
    try:
        # textAside = myInfo.split("|")[0]
        # inputContent = myInfo.split("|")[1]
        element = findEleByDetail(driver, "xpath", "//strong[.="+textAside+"]/following-sibling::input")
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        raise e

def selectValueByTextAside(myInfo,*arg):       # 根据输入框旁边的字段定位并向下拉框输值,待整理参数，TODO
    global driver
    try:
        textAside = myInfo.split("|")[0]
        inputContent = myInfo.split("|")[1]
        element = Select(findEleByDetail(driver, "xpath", "//strong[.="+textAside+"]/following-sibling::select"))
        element.select_by_visible_text(inputContent)
    except Exception as e:
        raise e

def ifExistThenClick(locationType,locatorExpression,*arg):     # 若元素存在，则点击
    try:
        # element = findEleByDetail(driver,locationType,locatorExpression)
        element = WebDriverWait(driver, 5).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        element.click()
    except Exception as e:
        pass

def ifExistThenSendkeys(locationType,locatorExpression,inputContent):     # 若元素存在，则输值
    try:
        # element = findEleByDetail(driver,locationType,locatorExpression)
        element = WebDriverWait(driver, 5).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        pass

def setDataByJS(locationType,locatorExpression,inputContent):       # 通过js修改日期空间的“readonly属性”
    try:
        element = findEleByDetail(driver,locationType,locatorExpression)
        removeAttribute(driver,element,"readonly")
        element.send_keys(inputContent)
    except Exception as e:
        pass

def BoxHandler(locationType,locatorExpression,textInBox):       # 若存在弹出框，则处理点击
    try:
        sleep(1)
        assert_string_in_pagesourse(textInBox)
        click_Obj(locationType,locatorExpression)
    except Exception as e:
        pass