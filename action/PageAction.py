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


'''
【关键字分类】
1、浏览器操作:open_browser、visit_url、close_browser、close_page、switch_to_frame、switch_to_default_content、
            maximize_browser、switch_to_now_window、refresh_page、scroll_slide_field；
2、常规操作：clear、specObjClear、click_Obj、click_SpecObj、sendkeys_To_Obj、sendkeys_To_SpecObj、SelectValues、
    xpath_combination_click、xpath_combination_click_loop、xpath_combination_send_keys、xpath_combination_click_send_keys_loop、
    menu_select、capture_screen（setValueByTextAside、selectValueByTextAside,capture_screen_old）；
3、辅助定位：highlightElement、highlightElements、whichIsEnabled、whichIsDisplayed；
4、获取信息：getTitle、getPageSource、getAttribute、getDate_Now；
5、断言及判断：assert_string_in_pagesourse、assert_title、assert_list；
6、剪贴板操作：paste_string、press_key；
7、等待：loadPage、sleep、waitPresenceOfElementLocated、waitVisibilityOfElementLocated、wait_elements_vanish
        waitFrameToBeAvailableAndSwitchToIt；
8、鼠标键盘模拟：moveToElement、init_Mouse、pageKeySimulate；
9、外部程序调用：runProcessFile、page_upload_file（uploadFile_x1、uploadFile_x2）；
10、字符串操作：randomNum、pinyinTransform、compose_JSON；
11、带判断关键字：ifExistThenClick、ifExistThenSendkeys、BoxHandler、ifExistThenSelect、ifExistThenSetData、ifExistThenReturnAttribute_pinyin、
    ifExistThenReturnOperateValue、ifExistThenChooseOperateValue、ifExistThenChooseOperateValue_diffPosition、
    ifExistThenPass_xpath_combination
12、JS相关：setDataByJS；
13、项目关键字：销售合同新增+审批：finalBoxClick、ifDoubleMsg（writeContracNum）
               项目关键字：采购模块：checkApprover
'''
# ****************************************浏览器操作****************************************

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
        if url == u'200':
            driver.get('http://kdevelop.chinacloudapp.cn:9002/login.html')
        elif url == u'500':
            driver.get('http://kintergration.chinacloudapp.cn:9002/login.html')
        elif url == u'700':
            driver.get('http://kdevelop.chinacloudapp.cn:9003/login.html')
        elif url == u'810':
            driver.get('http://d58d71fb-74bb-470d-b173-a2f6ca23f9c2.chinacloudapp.cn')
        else:
            driver.get(url)
    except Exception as e:
        raise e

def close_browser(*arg):        #关闭浏览器
    global driver
    try:
        driver.quit()
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

def maximize_browser():     #窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception as e:
        raise e

def switch_to_now_window(handlesNum,*arg):      #切换进入frame
    global driver
    try:
        handlesNum = int(handlesNum)
        all_handles = driver.window_handles
        driver.switch_to.window(all_handles[handlesNum])
        print(all_handles)
    except Exception as e:
        print('未找到指定句柄')
        raise e

def close_page(*arg):  # 关闭标签页
    global driver
    try:
        driver.close()
    except Exception as e:
        raise e

def refresh_page(*arg):        #刷新网页
    global driver
    try:
        driver.refresh()
    except Exception as e:
        raise e

# 滚动条上下移动，拖动到可见的元素去
def scroll_slide_field(locationType, locatorExpression, *arg):
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        driver.execute_script("arguments[0].scrollIntoView();", element)  # 拖动到可见的元素去
    except Exception as e:
        raise e

# ****************************************常规操作****************************************

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

def SelectValues(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    try:
        el = Select(findEleByDetail(driver,locationType,locatorExpression))
        el.select_by_visible_text(inputContent)
    except Exception as e:
        raise e

def xpath_combination_click(attributeType, locatorExpression, attributeValue, *arg):
    # 将“操作值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“操作值”放入“元素定位表达式”的“[]”的指定属性值中，由xpath定位元素后，并执行点击操作
    try:
        combination_left = locatorExpression.split("[]")[0]
        combination_right = locatorExpression.split("[]")[1]
        if attributeType == "text()":
            combination = combination_left + '[' + attributeType +'="' + attributeValue + '"]' + combination_right
        else:
            combination = combination_left + '[@' + attributeType +'="' + attributeValue + '"]' + combination_right
        element = findElebyMethod(driver, 'xpath', combination)
        element.click()
    except Exception as e:
        raise e

def xpath_combination_click_loop(attributeType, locatorExpression, attributeValues, *arg):
    # 操作值格式:属性值|属性值|属性值... 根据属性值数量，循环点击操作
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行点击操作
    try:
        loop_time = attributeValues.count("|") + 1
        attributeValue = attributeValues.split("|")
        # 循环
        for i in range(loop_time):
            xpath_combination_click(attributeType, locatorExpression, attributeValue[i])
    except Exception as e:
        raise e

def xpath_combination_send_keys(attributeType, locatorExpression, attributeValue_sendValue, *arg):
    # 操作值格式：属性值|输入值
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行输入操作
    try:
        attributeValue = attributeValue_sendValue.split("|")[0]
        sendValue = attributeValue_sendValue.split("|")[1]
        # 拼接出指定Xpath
        combination_left = locatorExpression.split("[]")[0]
        combination_right = locatorExpression.split("[]")[1]
        if attributeType == "text()":
            combination = combination_left + '[' + attributeType +'="' + attributeValue + '"]' + combination_right
        else:
            combination = combination_left + '[@' + attributeType +'="' + attributeValue + '"]' + combination_right
        # 依据拼接的Xpath，查找指定元素
        element = findElebyMethod(driver, 'xpath', combination)
        element.send_keys(sendValue)
    except Exception as e:
        raise e

def xpath_combination_send_keys_loop(attributeType, locatorExpression, attributeValues_sendValues, *arg):
    # 操作值格式：属性值|输入值[]属性值|输入值... 根据属性值数量，循环输入操作
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行输入操作
    try:
        loop_time = attributeValues_sendValues.count("[]") + 1
        attributeValue_sendValue = attributeValues_sendValues.split("[]")
        # 循环
        for i in range(loop_time):
            xpath_combination_send_keys(attributeType, locatorExpression, attributeValue_sendValue[i])
    except Exception as e:
        raise e

def menu_select(menu_text,*arg):
    # 操作值格式：模块名称|菜单名称|菜单选项
    # 打开此菜单页面
    try:
        select_time = menu_text.count("|")
        menu_operation = menu_text.split("|")
        # 判断
        if select_time == 1:
            # 鼠标移动到模块名称上
            xpath_1 = '//ul[@class="navlist"]/li/a[contains(text(),"' + menu_operation[0] + '")]'
            moveToElement('xpath', xpath_1)
            sleep(0.5)
            # 鼠标点击菜单名称
            xpath_2 = '//ul[contains(@class,"hasSubMore")]/li/a[text()="' + menu_operation[1] + '"]'
            click_Obj('xpath', xpath_2)
        elif select_time == 2:
            # 鼠标移动到模块名称上
            xpath_1 = '//ul[@class="navlist"]/li/a[contains(text(),"' + menu_operation[0] + '")]'
            moveToElement('xpath', xpath_1)
            sleep(0.5)
            # 鼠标移动到菜单名称上
            xpath_2 = '//ul[contains(@class,"hasSubMore")]/li/a[text()="' + menu_operation[1] + '"]'
            moveToElement('xpath', xpath_2)
            sleep(0.5)
            # 鼠标点击菜单名称
            xpath_3 = '//ul[contains(@class,"subnavlist2")]/li/a[text()="' + menu_operation[2] + '"]'
            click_Obj('xpath', xpath_3)
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

def capture_screen_old(*arg):       #截图，旧，该方法在日期路径下无法区分具体流程的截图
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

def capture_screen(picDir,*arg):       #截图，新，保存截图路径从外部传进来，可在一级目录下添加二级目录
    global driver
    # 获取当前时间，精确到秒
    currentTime = getCurrentTime()
    # 拼接一场图片保存的绝对路径及名称
    picNameAndPath = str(picDir) + "\\" + str(currentTime) + ".png"
    try:
        # 截屏，并保存为本地图片
        driver.get_screenshot_as_file(picNameAndPath.replace('\\',r'\\'))
        # print("picNameAndPath 为：",picNameAndPath.replace('\\',r'\\'))
    except Exception as e:
        raise e
    else:
        return picNameAndPath



# ****************************************辅助定位****************************************

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

# ****************************************获取信息****************************************

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

def getAttribute(locationType,locatorExpression,attributeType,*arg):        # 获取页面元素属性值
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        attributeValue = element.get_attribute(attributeType)
        if attributeValue is None:
            if attributeType == "text":
                attributeValue = element.text
        return attributeValue
    except Exception as e:
        raise e


def getDate_Now(MyStr,*arg):        # 获取指定连接符的当前日期，20180517
    try:
        import datetime
        MyDate = datetime.datetime.now().strftime("%Y"+MyStr+"%m"+MyStr+"%d")
        print('********** 返回日期为：',MyDate,' **********')
        return MyDate
    except Exception as e:
        raise e

# ****************************************断言及判断****************************************

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

def assert_title(titleStr,*arg):     #断言判断当前页面标题是否存在指定字段
    global driver
    try:
        assert titleStr in driver.title, \
            u"当前不存在标题为 %s 的页面" % titleStr
    except AssertionError as e:
        raise AssertionError(e)
    except Exception as e:
        raise e

def assert_list(locationType, locatorExpression, listStr, *arg):    # 断言判断下拉框选项是否包含指定值
    '''若校验多值，用“|”分开'''
    global driver
    element = findElebyMethod(driver, locationType, locatorExpression)
    # 获取下拉框所有元素对象
    myOptions = Select(element).options
    myOptionValues = map(lambda option: option.text, myOptions)
    # print("myOptionValues:",myOptionValues)

    for option in listStr.split("|"):
        try:
            assert option in myOptionValues, u"下拉框中未找到该元素: %s ！" % option
        except AssertionError as e:
            raise AssertionError(e)
        except Exception as e:
            raise e

# ****************************************剪贴板操作****************************************

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


# ****************************************等待****************************************

def loadPage(loop_time=10,*arg):     # 设置页面加载时间
    global driver
    try:
        sleep(0.5)
        driver.set_page_load_timeout(10)
        # 等待加载动图消失
        wait_elements_vanish('xpath','//div[@id="loading" and contains(@style,"display: block;")]',loop_time)
    except TimeoutError as e:
        print("********** 等待页面加载超时 **********")
        raise TimeoutError(e)

def sleep(sleepSeconds,*arg):       #强制等待
    try:
        time.sleep(float(sleepSeconds))
    except Exception as e:
        raise e

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

def wait_elements_vanish(locationType,locatorExpression,loop_time=10,*arg):
    # 等待指定元素从页面中消失后，再进行下一步
    global driver
    driver.implicitly_wait(0)
    for i in range(int(loop_time)):
        try:
            time.sleep(1)
            elements = driver.find_elements(by = locationType, value = locatorExpression)
            if not elements or not elements[0].is_displayed():
                return True
        except:
            return True
    driver.implicitly_wait(1)
    from selenium.common.exceptions import TimeoutException
    raise TimeoutException

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


# ****************************************鼠标键盘模拟****************************************

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
        if keyType == "page_tab":
            element.send_keys(Keys.TAB)
        if keyType == "page_left":
            element.send_keys(Keys.LEFT)
        if keyType == "page_right":
            element.send_keys(Keys.RIGHT)
    except Exception as e:
        raise e

# ****************************************外部程序调用****************************************

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
        print("********** 调用文件的绝对路径为：",myPath," **********")
        os.system(myPath)
        sleep(1)
    except Exception as e:
        raise e

def page_upload_file(locationType,locatorExpression,uploadFileName, *arg):
    # 先点击上传按钮，再运行autoit文件
    # autoit上传文件，使用相对路径，需在操作值内输入上传文件名
    # 上传文件放置目录：..\fileHandle\upload_file
    global driver
    try:
        findElebyMethod(driver, locationType, locatorExpression).click()
        filePath = parentDirPath + u"\\fileHandle\\" + "file_upload_script.exe"
        print("********** 调用文件的绝对路径为：", filePath, " **********")
        uploadPath = parentDirPath + u"\\fileHandle\\upload_file\\" + uploadFileName
        print("********** 上传文件的绝对路径为：", uploadPath, " **********")
        cmd = "%s %s" %(filePath ,uploadPath)
        os.popen(cmd)
        sleep(3)
    except Exception as e:
        raise e

# ****************************************字符串操作****************************************

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

def pinyinTransform(myStr,*arg):        # 将汉字转换成拼音
    try:
        import pypinyin
        from pypinyin import pinyin, lazy_pinyin
        strTransformed = ''.join(lazy_pinyin(myStr))
        return strTransformed
    except Exception as e:
        raise e

def compose_JSON(coordinate, content):
    try:
        # coordinate、content以"[]"为分割符
        # 以{coordinate.split("[]")[i]:content.split("[]")[i]}，组成JSON
        JSON_return = {}
        for i in range(len(coordinate.split("[]"))):
            JSON_k = coordinate.split("[]")[i]
            JSON_v = content.split("[]")[i]
            JSON_return[JSON_k] = JSON_v
        return JSON_return
    except Exception as e:
        raise e

# ****************************************带判断关键字****************************************

def ifExistThenClick(locationType,locatorExpression,*arg):     # 若元素存在，则点击
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        element.click()
    except Exception as e:
        pass

def ifExistThenSendkeys(locationType,locatorExpression,inputContent):     # 若元素存在，则输值
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        element.clear()
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

def ifExistThenSelect(locationType,locatorExpression,inputContent):     # 若元素存在，则选择选项
    global driver
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        el = Select(element)
        el.select_by_visible_text(inputContent)

    except Exception as e:
        pass

def ifExistThenSetData(locationType,locatorExpression,inputContent):
    # 若元素存在，则输入日期
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        removeAttribute(driver,element,"readonly")
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        pass

def ifExistThenReturnAttribute_pinyin(locationType,locatorExpression,attributeType,*arg):
    # 若元素存在，则获取页面元素属性值，并转化为拼音字母（查审批岗位专用）
    global driver
    from pypinyin import lazy_pinyin
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        attributeValue = element.get_attribute(attributeType)
        if attributeValue is None:
            if attributeType == "text":
                attributeValue = element.text
        # 将attributeValue以逗号分割，用第一个值转为拼音
        if "," in attributeValue:
            attributeValue = attributeValue.split(",")[0]
        # 将属性值转为拼音字母
        strTransformed = ''.join(lazy_pinyin(attributeValue))

        return strTransformed
    except Exception as e:
        return ""

def ifExistThenReturnOperateValue(locationType, locatorExpression, operateValue, *arg):
    # 若元素存在，则返回表格操作值
    global driver
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None:
            return operateValue
    except Exception as e:
        return ""

def ifExistThenChooseOperateValue(locationType, locatorExpression, operateValue, *arg):
    # 返回值格，需填写一个位置信息。两个返回值择一，填入同一个格中。
    # 表格操作值填写格式：元素存在时返回值|元素不存在时返回值
    # 若元素存在，则返回表格操作值中，"|"之前的值；元素不存在，则返回"|"之后的值
    global driver
    try:
        driver.implicitly_wait(1)
        exist_value = operateValue.split("|")[0]
        not_exist_value = operateValue.split("|")[1]
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None:
            return exist_value
    except Exception as e:
        return not_exist_value

def ifExistThenChooseOperateValue_diffPosition(locationType, locatorExpression, operateValue, *arg):
    # 返回值格，需填写两个位置信息，中间以"[]"分隔。两个返回值择一，填入不同格中。
    # 表格操作值填写格式：元素存在时返回值|元素不存在时返回值
    # 若元素存在，则返回表格操作值中，"|"之前的值，写入"[]"之前的坐标中；
    # 元素不存在，则返回"|"之后的值，写入"[]"之后的坐标中；
    global driver
    try:
        driver.implicitly_wait(1)
        # 元素存在时，返回的值
        exist_value = operateValue.split("|")[0]
        exist_return_value = "%s[]" %(exist_value)
        # 元素不存在时，返回的值
        not_exist_value = operateValue.split("|")[1]
        not_exist_return_value = "[]%s" % (not_exist_value)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None:
            return exist_return_value
    except Exception as e:
        return not_exist_return_value

def ifExistThenPass_xpath_combination(attributeType, locatorExpression, attributeValue, *arg):
    # 将“操作值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“操作值”放入“元素定位表达式”的“[]”的指定属性值中，由xpath定位元素是否存在，存在则通过，不存在则报错
    global driver
    try:
        driver.implicitly_wait(1)
        # 拼接Xpath
        combination_left = locatorExpression.split("[]")[0]
        combination_right = locatorExpression.split("[]")[1]
        if attributeType == "text()":
            combination = combination_left + '[' + attributeType +'="' + attributeValue + '"]' + combination_right
        else:
            combination = combination_left + '[@' + attributeType +'="' + attributeValue + '"]' + combination_right
        # 由xpath定位元素是否存在
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by='xpath', value=combination))
        assert element
    except Exception as e:
        raise e

# ****************************************JS相关****************************************

def setDataByJS(locationType,locatorExpression,inputContent):       # 通过js修改日期空间的“readonly属性”
    try:
        element = findEleByDetail(driver,locationType,locatorExpression)
        removeAttribute(driver,element,"readonly")
        element.clear()
        input_time = inputContent.split(" ")[0]
        element.send_keys(input_time)
    except Exception as e:
        pass


# ****************************************项目关键字：销售合同新增+审批****************************************

def writeContracNum(myInfo,*arg):
    # 该方法加断点时可往excel中写值成功，不加断点则写不进去，randomContracNum 方法暂时启用，TODO
    try:
        # ParseExcel().randomContracNum(myInfo)
        randContractNum = myInfo + randomNum(9)
        return randContractNum
    except Exception as e:
        raise e

def finalBoxClick(*arg):        # 处理合同审批后，弹出窗口点击操作（双判断效率较低）,20180517
    global driver
    try:
        # element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = "xpath",
        #                                                                   value = "//button[.=\"返回销售合同待办\"]"))
        element = driver.find_element_by_xpath("//button[.=\"返回销售合同待办\"]")
        element.click()
    except Exception as e:
        # element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = "xpath",
        #                                                                   value = "//button[.=\"返回我的单据\"]"))
        element = driver.find_element_by_xpath("//button[.=\"返回我的单据\"]")
        element.click()

def ifDoubleMsg(locationType,locatorExpression,*arg):      # 在销售合同新增-文本信息中，判断是否只有一条收款条款
    global driver
    try:
        myValue = getAttribute(locationType,locatorExpression,"value")
        if myValue == "":
            click_Obj("xpath","//table[@id=\"sktkList\"]/descendant::tr[3]/td[6]/a")
        else:
            pass
    except Exception as e:
        raise e

# ****************************************项目关键字：采购模块****************************************
def checkApprover(varInfo):
    '''
    验证如果info1=info2时，info3是否包含于info4，用于判断审批角色是否正确
    :param varInfo:格式为“info1|info2|info3|info4”，info内部用“、”隔开
    :return:None
    '''
    myCount = varInfo.count("|")
    conditionNum = int((myCount-1)/2)

    JSON_Info = {}
    for i in range(1,conditionNum+1):
        info_Key = varInfo.split("|")[2*i-2]
        info_Value = varInfo.split("|")[2*i-1]
        JSON_Info[info_Key] = info_Value
        try:
            assert info_Key in info_Value
        except:
            print("********** 该流程不符合此条件校验 **********")
            print(JSON_Info)
            return

        # 根据业务规则校验审批人范围是否正确
        myValue = varInfo.split("|")[myCount - 1]
        approverNames = varInfo.split("|")[myCount]
        myNames = approverNames.split("、")

        try:
            assert myValue in myNames, u"下一岗审批人为'%s'，与预期不符 ！" %myValue
        except AssertionError as e:
            raise AssertionError(e)
        except Exception as e:
            raise e