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
# from selenium.webdriver.firefox.options import Options

# 定义全局driver变量
driver = None
# 全局的等待类实例对象
waitUtil = None


'''
【关键字分类】
[报错类]
CNBMError（class）、CNBMException.

[KER]
1、浏览器操作:open_browser、visit_url、close_browser、close_page、switch_to_frame、switch_to_default_content、
            maximize_browser、switch_to_now_window、refresh_page、scroll_slide_field；
2、常规操作：clear、specObjClear、click_Obj、click_SpecObj、sendkeys_To_Obj、sendkeys_To_SpecObj、sendkeys_to_elements、SelectValues、
    xpath_combination_click、xpath_combination_click_loop、xpath_combination_send_keys、xpath_combination_click_send_keys_loop、
    xpath_combination_send_keys_click_loop、menu_select、
    capture_screen（setValueByTextAside、selectValueByTextAside,capture_screen_old）；
3、辅助定位：highlightElement、highlightElements、whichIsEnabled、whichIsDisplayed；
4、获取信息：getTitle、getPageSource、getAttribute、getDate_Now、getDateCalcuated、getTextInTable；
5、断言及判断：assert_string_in_pagesourse、assert_title、assert_list；
6、剪贴板操作：paste_string、press_key；
7、等待：loadPage、sleep、waitPresenceOfElementLocated、waitVisibilityOfElementLocated、wait_elements_vanish
        waitFrameToBeAvailableAndSwitchToIt；
8、鼠标键盘模拟：moveToElement、init_Mouse、pageKeySimulate、get_clipboard_return；
9、外部程序调用：runProcessFile、page_upload_file；
10、字符串操作：randomNum、pinyinTransform、compose_JSON；
11、带判断关键字：ifExistThenClick、ifExistThenSendkeys、BoxHandler、ifExistThenSelect、ifExistThenSetData、ifExistThenReturnAttribute_pinyin、
    ifExistThenReturnOperateValue、ifExistThenChooseOperateValue、ifExistThenChooseOperateValue_diffPosition、
    ifExistThenPass_xpath_combination
12、JS相关：setDataByJS；
13、项目关键字：销售合同新增+审批：finalBoxClick、ifDoubleMsg（writeContracNum）
                采购模块：checkApprover；
                财务管理模块：getHidenInfo、getInfoWanted、getNextUser（适用于所有审批）、setAmountOfPayment、
                            ifExist_pageKeySimulate；
                进出口合同模块：getNumWanted；
                日常办公：getInfoNeeded;
                组合功能：getApprovalFlow、loginProcess；
                
[SAP]
1、基础配置：createObject、saplogin、updateActiveWindow、closeSAP、createNewSession、closeAllSession；
2、基本操作：performById、、getObj、getText、getNumInText；
3、等待及校验：waitObj、waitUntil、checkText；
4、项目关键字：
        采购申请：chooseToReturn、chooseHowToTrans；
'''



# ****************************************浏览器操作****************************************
class CNBMError(Exception):
    def __init__(self,ErrorInfo):
        super().__init__(self) #初始化父类
        self.errorinfo=ErrorInfo

    def __str__(self):
        return self.errorinfo

def SAPException(func):
    def wrapper(*args, **kwargs):
        try:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # closeSAP()            # 关闭所有打开的sap进程
                # closeAllSession()       # 关闭本流程的所有sap进程
                funcName = func.__name__
                errInfo = "\n[关键字] " + funcName \
                          + "\n[异常信息] %s" %repr(e)
                raise CNBMError(errInfo)
        except CNBMError as err:
            raise err
    return wrapper

# ****************************************浏览器操作****************************************

def open_browser(browserName,*arg):        #打开浏览器
    global driver,waitUtil
    try:
        if browserName.lower() == 'ie':
            driver = webdriver.Ie()
        elif browserName.lower() == 'chrome':
            chrome_options = webdriver.ChromeOptions()
            # 用于控制进程是否在后台执行
            # chrome_options.add_argument('--headless')
            # chrome_options.add_argument('--disable-gpu')
            driver = webdriver.Chrome(chrome_options=chrome_options)
        elif browserName.lower() == 'edge':
            driver = webdriver.Edge()
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
            driver.get('http://cdwpdev01.chinacloudapp.cn:9001/login.html')
        elif url == u'400':
            driver.get('http://cdwpdev01.chinacloudapp.cn:9004/login.html')
        elif url == u'450':
            driver.get('http://cdwpdev01.chinacloudapp.cn:9003/login.html')
        elif url == u'500':
            driver.get('http://kintergration.chinacloudapp.cn:9002/login.html')
        elif url == u'510':
            driver.get('http://kintergration01.chinacloudapp.cn:9510/#/as/login')
        elif url == u'520':
            driver.get('http://kintergration01.chinacloudapp.cn:9520/#/as/login')
        elif url == u'530':
            driver.get('http://kintergration01.chinacloudapp.cn:9530/#/as/login')
        elif url == u'540':
            driver.get('http://kintergration01.chinacloudapp.cn:9540/#/as/login')
        elif url == u'600':
            driver.get('http://kintergration.chinacloudapp.cn:9003/login.html')
        elif url == u'700':
            driver.get('http://kdevelop.chinacloudapp.cn:9003/login.html')
        elif url == u'810':
            driver.get('http://pre-mongodb-01.chinacloudapp.cn:9003/#/as/login')
        elif url == u'800':
            # 线上
            driver.get('http://cdwp.cnbmxinyun.com/#/as/login')
        else:
            driver.get(url)
    except Exception as e:
        raise e

def close_browser(*arg):        #关闭浏览器
    global driver
    try:
        if driver:
            driver.quit()
            driver = None
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

def switch_to_now_window(handlesNum=0,*arg):      #切换进入frame
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

# 滚动条上下移动，滑动到可见的元素，并将其置顶，若操作值填down，则置于底部
def scroll_slide_field(locationType, locatorExpression, position="up", *arg):
    global driver
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        if position == "up":
            driver.execute_script("arguments[0].scrollIntoView();", element)  # 滑动到可见的元素，并将其置顶
        elif position == "down":
            driver.execute_script("arguments[0].scrollIntoView(false);", element)  # 滑动到可见的元素，并将其置底
    except Exception as e:
        raise e

def scroll_into_field(locationNum="0|0", *arg):
    global driver
    # 将浏览器页面滚动到指定位置，在操作值中输入横纵坐标数值，操作值输入格式：x|y
    try:
        x = locationNum.split("|")[0]
        y = locationNum.split("|")[1]
        execute_string = "scrollTo(%s,%s);" %(x,y)
        driver.execute_script(execute_string)
        sleep(1.5)
    except Exception as e:
        raise e

# ****************************************常规操作****************************************

def clear(locationType,locatorExpression,*arg):     #清除输入框默认内容
    global driver
    flag = False
    try:
        el_init = findEleByDetail(driver,locationType,locatorExpression)
        flag = True
        el_init.clear()
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, el_init)
        raise e

def sendkeys_To_Obj(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    flag = False
    try:
        element = findEleByDetail(driver,locationType,locatorExpression)
        flag = True
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

def click_Obj(locationType, locatorExpression, *arg):       #点击页面元素
    global driver
    flag = False
    try:
        element = findEleByDetail(driver, locationType, locatorExpression)
        flag = True
        element.click()
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

# 针对partial_link_text、link_text、css_selector报错Unsupported locator strategy封装单独关键字

def specObjClear(locationType,locatorExpression,*arg):     #清除输入框默认内容，暂时弃用
    global driver
    flag = False
    try:
        element = findElebyMethod(driver,locationType,locatorExpression)
        flag = True
        element.clear()
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

def sendkeys_To_SpecObj(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    flag = False
    try:
        element = findElebyMethod(driver,locationType,locatorExpression)
        flag = True
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

def sendkeys_to_elements(locationType,locatorExpression,inputContent):
    # 查找多个元素，并向查找到的所有输入框中输入操作值。
    # 操作值可用“|”作为分隔符，将多个值依照查找顺序填入多个输入框中
    # 若操作值无分隔符，则所有输入框都输入同一操作值
    global driver
    flag = False
    try:
        elements = findElesbyMethod(driver, locationType, locatorExpression)
        flag = True
        if "|" in inputContent:
            inputArray = inputContent.split("|")
            loop_time = inputContent.count("|") + 1
            for i in range(loop_time):
                elements[i].clear()
                elements[i].send_keys(inputArray[i])
        else:
            for ele in elements:
                ele.clear()
                ele.send_keys(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlightElements(locationType, locatorExpression)
        raise e

def click_SpecObj(locationType, locatorExpression, *arg):       #点击页面元素
    global driver
    flag = False
    try:
        element = findElebyMethod(driver, locationType, locatorExpression)
        flag = True
        element.click()
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

def SelectValues(locationType,locatorExpression,inputContent):      #输入框输值
    global driver
    flag = False
    try:
        el_init = findEleByDetail(driver,locationType,locatorExpression)
        flag = True
        el = Select(el_init)
        el.select_by_visible_text(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, el_init)
        raise e

def xpath_combination_click(attributeType, locatorExpression, attributeValue, *arg):
    # 将“操作值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“操作值”放入“元素定位表达式”的“[]”的指定属性值中，由xpath定位元素后，并执行点击操作
    # “元素定位方式”中填入HTML属性，如text()、id、class，以“starts-with(”开头是xpath里starts-with()的意思
    '''
    例：元素定位方式：starts-with(text()
    元素定位表达式：//*[]/../td[4]/input
    操作值：6%
    得出的Xpath：//*[starts-with(text(),"6%")]/../td[4]/input
    '''
    try:
        combination_left = locatorExpression.split("[]")[0]
        combination_right = locatorExpression.split("[]")[1]

        if attributeType.startswith("starts-with("):
            attributeType = attributeType[12:]
            if attributeType == "text()":
                combination = combination_left + '[starts-with(' + attributeType + ',"' + attributeValue + '")]' + combination_right
            else:
                combination = combination_left + '[starts-with(@' + attributeType + ',"' + attributeValue + '")]' + combination_right
        elif attributeType == "text()":
            combination = combination_left + '[' + attributeType +'="' + attributeValue + '"]' + combination_right
        else:
            combination = combination_left + '[@' + attributeType +'="' + attributeValue + '"]' + combination_right
        highlightElement('xpath', combination)
        click_Obj('xpath', combination)
    except Exception as e:
        raise e

def xpath_combination_click_loop(attributeType, locatorExpression, attributeValues, *arg):
    # 操作值格式:属性值|属性值|属性值... 根据属性值数量，循环点击操作，英文逗号也可做分隔符（不常用）
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行点击操作
    try:
        # '|'作为分隔符
        if '|' in attributeValues:
            loop_time = attributeValues.count("|") + 1
            attributeValue = attributeValues.split("|")
            # 循环
            for i in range(loop_time):
                xpath_combination_click(attributeType, locatorExpression, attributeValue[i])
        # ','作为分隔符
        elif ',' in attributeValues:
            loop_time = attributeValues.count(",") + 1
            attributeValue = attributeValues.split(",")
            # 循环
            for i in range(loop_time):
                xpath_combination_click(attributeType, locatorExpression, attributeValue[i])
        else:
            xpath_combination_click(attributeType, locatorExpression, attributeValues)
    except Exception as e:
        raise e

def xpath_combination_send_keys(attributeType, locatorExpression, attributeValue_sendValue, *arg):
    # 操作值格式：属性值|输入值
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行输入操作
    # “元素定位方式”中填入HTML属性，如text()、id、class，以“starts-with(”开头是xpath里starts-with()的意思
    '''
    例：元素定位方式：starts-with(text()
    元素定位表达式：//*[]/../td[4]/input
    操作值：6%|170
    得出的Xpath：//*[starts-with(text(),"6%")]/../td[4]/input
    往此Xpath元素中输入：170
    '''
    try:
        attributeValue = attributeValue_sendValue.split("|")[0]
        sendValue = attributeValue_sendValue.split("|")[1]
        # 拼接出指定Xpath
        combination_left = locatorExpression.split("[]")[0]
        combination_right = locatorExpression.split("[]")[1]

        if attributeType.startswith("starts-with("):
            attributeType = attributeType[12:]
            if attributeType == "text()":
                combination = combination_left + '[starts-with(' + attributeType + ',"' + attributeValue + '")]' + combination_right
            else:
                combination = combination_left + '[starts-with(@' + attributeType + ',"' + attributeValue + '")]' + combination_right
        elif attributeType == "text()":
            combination = combination_left + '[' + attributeType +'="' + attributeValue + '"]' + combination_right
        else:
            combination = combination_left + '[@' + attributeType +'="' + attributeValue + '"]' + combination_right
        # 依据拼接的Xpath，查找指定元素，执行输入操作
        sendkeys_To_Obj('xpath', combination, sendValue)
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

def xpath_combination_send_keys_click_loop(attributeType, locatorExpression, attributeValues_sendValues, *arg):
    # 操作值格式：属性值|输入值[]属性值|输入值... 根据属性值数量，循环输入操作、点击操作
    # 将“属性值”与“元素定位表达式”拼接到一起组成完整表达式定位元素
    # 将“属性值”放入“元素定位表达式”的“[]”的指定属性中，由xpath定位元素后，并执行输入操作、点击操作
    try:
        # 操作值处理
        attributeValues_click = ''
        attributeValues_sendValues_send_keys = ''
        attributeValue_sendValue_array = attributeValues_sendValues.split("[]")
        # json格式：{物料内部编码[i]: 数量[i]}
        JSON_A = {}
        for i in range(len(attributeValue_sendValue_array)):
            JSON_k = attributeValue_sendValue_array[i].split("|")[0]
            JSON_v = attributeValue_sendValue_array[i].split("|")[1]
            JSON_A[JSON_k] = JSON_v
        # 输入操作值
        for i in range(len(attributeValue_sendValue_array)):
            if attributeValue_sendValue_array[i].endswith('|'):
                continue
            attributeValues_sendValues_send_keys = attributeValues_sendValues_send_keys + '[]' + attributeValue_sendValue_array[i]
        attributeValues_sendValues_send_keys = attributeValues_sendValues_send_keys[2:]
        # 点击操作值
        for (k, v) in JSON_A.items():
            attributeValues_click = attributeValues_click + '|' + k
        attributeValues_click = attributeValues_click[1:]
        # 定位表达式处理
        # 例：表达式为//td[text()="02230HBV-F30"]/../td[1]/span/../..//input[contains(@name,"referenceCount")]
        # 结果为//td[text()="02230HBV-F30"]/../td[1]/span
        str_position = locatorExpression.rindex('/../../')
        locatorExpression_click = locatorExpression[:str_position]
        # 调用循环函数
        xpath_combination_send_keys_loop(attributeType, locatorExpression, attributeValues_sendValues_send_keys)
        xpath_combination_click_loop(attributeType, locatorExpression_click, attributeValues_click)
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
            xpath_1 = '//span[@class="nav-text" and text()="' + menu_operation[0] + '"]/..'
            moveToElement('xpath', xpath_1)
            sleep(0.5)
            # 鼠标点击菜单名称
            xpath_2 = '//span[@class="nav-text" and text()="' + menu_operation[0] + '"]/../following-sibling::*[1]/li/a/span[text()="' + menu_operation[1] + '"]/..'
            click_Obj('xpath', xpath_2)
        elif select_time == 2:
            # 鼠标移动到模块名称上
            xpath_1 = '//span[@class="nav-text" and text()="' + menu_operation[0] + '"]/..'
            moveToElement('xpath', xpath_1)
            sleep(0.5)
            # 鼠标点击菜单名称
            xpath_2 = '//span[@class="nav-text" and text()="' + menu_operation[0] + '"]/../following-sibling::*[1]/li/a/span[text()="' + menu_operation[1] + '"]/..'
            click_Obj('xpath', xpath_2)
            sleep(0.5)
            # 鼠标点击菜单名称
            xpath_3 = '//span[@class="nav-text" and text()="' + menu_operation[1] + '"]/../following-sibling::*[1]/li/a/span[text()="' + menu_operation[2] + '"]/..'
            click_Obj('xpath', xpath_3)
    except Exception as e:
        raise e

def setValueByTextAside(textAside,inputContent,*arg):       # 根据输入框旁边的字段定位并向输入框输值,待整理参数，TODO
    global driver
    flag = False
    try:
        # textAside = myInfo.split("|")[0]
        # inputContent = myInfo.split("|")[1]
        element = findEleByDetail(driver, "xpath", "//strong[.="+textAside+"]/following-sibling::input")
        flag = True
        element.clear()
        element.send_keys(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
        raise e

def selectValueByTextAside(myInfo,*arg):       # 根据输入框旁边的字段定位并向下拉框输值,待整理参数，TODO
    global driver
    flag = False
    try:
        textAside = myInfo.split("|")[0]
        inputContent = myInfo.split("|")[1]
        element = Select(findEleByDetail(driver, "xpath", "//strong[.="+textAside+"]/following-sibling::select"))
        flag = True
        element.select_by_visible_text(inputContent)
    except Exception as e:
        if flag:
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            # 找到元素，但后续失败时，可通过截图查看报错高亮元素
            highlight(driver, element)
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

def capture_screen(picDir, *arg):       #截图，新，保存截图路径从外部传进来，可在一级目录下添加二级目录
    import win32api, win32con
    from PIL import ImageGrab
    global driver, session
    # 获取当前时间，精确到秒
    currentTime = getCurrentTime()
    # 拼接一场图片保存的绝对路径及名称
    picNameAndPath = str(picDir) + "\\" + str(currentTime) + ".png"
    try:
        # 截屏，并保存为本地图片
        if driver:
            driver.get_screenshot_as_file(picNameAndPath.replace('\\',r'\\'))
        elif session:
            ''' 方法一：部分截图 '''
            # im = ImageGrab.grab()
            # im.save(picNameAndPath.replace('\\',r'\\'))
            ''' 方法二：全屏截图 '''
            win32api.keybd_event(win32con.VK_SNAPSHOT, 0)
            time.sleep(0.5)
            im=ImageGrab.grabclipboard()
            im.save(picNameAndPath.replace('\\',r'\\'))
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

def getDateCalcuated(MyStr):
    '''获取运算后的日期（陈卓/20190123）
    :param dimension: day/month，选择计算的颗粒度
    :param accuracy: 时间差，可为负
    :param hyphen: 日期连词符
    :return:
    '''
    try:
        import datetime
        from dateutil.relativedelta import relativedelta
        dimension = MyStr.split("|")[0]
        accuracy = int(MyStr.split("|")[1])
        hyphen = MyStr.split("|")[2]

        MyDate = datetime.datetime.today()
        if dimension == "month":
            MyDate = MyDate + relativedelta(months=accuracy)
        if dimension == "day":
            MyDate = MyDate + datetime.timedelta(days=accuracy)
        finalDate = MyDate.strftime("%Y"+hyphen+"%m"+hyphen+"%d")
        print('********** 返回日期为：',finalDate,' **********')
        return finalDate
    except Exception as e:
        raise e

def getTextInTable(locationType, locatorExpression, myInfo):
    '''从table指定行（a）和指定列（b）获取值
    :param locationType: 定位table的属性
    :param locatorExpression: 定位table的属性值
    :param myInfo: “a|b”（a：行表头；b：列表头）
    '''
    global driver
    try:
        table = findElebyMethod(driver,locationType,locatorExpression)
        trList = table.find_elements_by_tag_name("tr")
        thList = trList[0].find_elements_by_tag_name("th")

        rowText = myInfo.split("|")[0]
        colText = myInfo.split("|")[1]
        flag = False

        for i in range(len(thList)):
            # 确定列
            if thList[i].text == colText:
                colNum = i
                break
        assert "colNum" in locals().keys(), "table中未找到该列！"

        for row in trList:
            # 遍历每行，包括表头
            try:
                # 确定行
                assert rowText in row.text
                flag = True
            except:
                continue

            assert flag == True, "table中未找到该行！"
            tdList = row.find_elements_by_tag_name("td")
            value = tdList[colNum].text
            return value

    except AssertionError as e:
        raise AssertionError(e)
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

def calculateToCheck(var):
    '''计算交易前后回显值是否正确'''
    try:
        var_A = float(var.split("|")[0].replace(",",""))
        var_B = var.split("|")[1]
        var_C = float(var.split("|")[2].replace(",",""))
        var_R = float(var.split("|")[3].replace(",",""))
        errInfo = "交易前后计算输值与预期不符！"
        if var_B == "+":
            assert (var_A + var_C) == var_R, errInfo
        if var_B == "-":
            assert (var_A - var_C) == var_R, errInfo
        if var_B == "*":
            assert (var_A * var_C) == var_R, errInfo
        if var_B == "/":
            assert (var_A / var_C) == var_R, errInfo
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

def press_twoKey(keyA, keyB):       #模拟单按键，如： "tab"、"enter"
    try:
        KeyboardKeys.twoKeys(keyA, keyB)
    except Exception as e:
        raise e


# ****************************************等待****************************************

def loadPage(loop_time=60):     # 设置页面加载时间
    global driver
    try:
        sleep(0.5)
        driver.set_page_load_timeout(10)
        # 等待加载动图消失
        wait_elements_vanish('xpath', '//div[@id="loading" and contains(@style,"display: block;")]', loop_time)
        wait_elements_vanish('xpath', '//div[@id="loading"]/img', loop_time)
    except Exception as e:
        print("********** 等待页面加载超时 **********")
        raise e

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

def wait_elements_vanish(locationType, locatorExpression, loop_time=60):
    # 等待指定元素从页面中消失后，再进行下一步
    global driver
    driver.implicitly_wait(0)
    for i in range(100):
        try:
            # time.sleep(1)
            elements = driver.find_elements(by = locationType, value = locatorExpression)
            if not elements or not elements[0].is_displayed():
                return

            assert i <= int(loop_time), "等待其消失的元素存在于页面中超过 %d s！" %int(loop_time)
            sleep(1)
        except AssertionError as e:
            print("等待其消失的元素存在于页面中超过 %d s！" %int(loop_time))
            raise AssertionError(e)
        except:
            return

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
        if keyType == "page_enter":
            element.send_keys(Keys.ENTER)
    except Exception as e:
        raise e

def get_clipboard_return(locationType, locatorExpression, *arg):
    # 将输入框中的内容存入剪贴板中，作为函数返回值
    global driver
    try:
        # 点入输入框，并将输入框中的内容存入剪贴板中
        element = findElebyMethod(driver, locationType, locatorExpression)
        element.click()
        sleep(0.5)
        element.send_keys(Keys.CONTROL,'a')
        sleep(0.5)
        element.send_keys(Keys.CONTROL,'c')
        # 获取剪贴板中的内容，并将其作为返回值
        clipboard_text = Clipboard.getText()
        # 将bytes转为str
        clipboard_text = bytes.decode(clipboard_text)
        return clipboard_text
    except Exception as e:
        raise e

# ****************************************外部程序调用****************************************

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
        sleep(2)
        filePath = parentDirPath + u"\\fileHandle\\" + "file_upload_script.exe"
        print("********** 调用文件的绝对路径为：", filePath, " **********")
        uploadPath = parentDirPath + u"\\fileHandle\\upload_file\\" + uploadFileName
        print("********** 上传文件的绝对路径为：", uploadPath, " **********")
        cmd = "%s %s" %(filePath ,uploadPath)
        os.popen(cmd)
        sleep(2)
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
        input_time = inputContent.split(" ")[0]
        element.send_keys(input_time)
    except Exception as e:
        pass

def ifExistThenReturnAttribute_pinyin(locationType,locatorExpression,attributeType,*arg):
    # 若元素存在，则获取页面元素属性值，并转化为拼音字母（查审批岗位专用）
    global driver
    from pypinyin import lazy_pinyin
    from selenium.common.exceptions import TimeoutException
    # 当用户姓名中有多音字或其他原因，会导致转换的拼音字符串与用户账号不对应。
    # 针对于此情况，函数会先从special_list字典中，检索用户对应的拼音字符串。若检索不到，则进行拼音转换。
    special_list = {
        u'李浩': 'lihao01',
        u'曾蓉琴': 'zengrongqin',
        u'朴贤国': 'piaoxianguo',
        u'王洋': 'wangyang01',
        u'张鑫': 'zhangxin01',
        u'曾春娇': 'zengchunjiao'
    }
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
        # 先从字典中检索用户对应的拼音字符串，若检索不到，则进行拼音转换
        for (k, v) in special_list.items():
            if attributeValue == k:
                return v
        strTransformed = ''.join(lazy_pinyin(attributeValue))
        return strTransformed
    except TimeoutException as e:
        return ""
    except Exception as e:
        raise e

def ifExistThenReturnStopFlag(locationType, locatorExpression, attributeValue, *arg):
    # 若元素存在，则返回表格操作值
    global driver
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None \
                and element.get_attribute("title") == attributeValue:
            return ""
    except Exception as e:
        return ""

def ifExistThenChooseOperateValue(locationType, locatorExpression, operateValue, *arg):
    # 返回值格，需填写一个位置信息。两个返回值择一，填入同一个格中。
    # 表格操作值填写格式：元素存在时返回值|元素不存在时返回值
    # 若元素存在，则返回表格操作值中，"|"之前的值；元素不存在，则返回"|"之后的值
    global driver
    from selenium.common.exceptions import TimeoutException
    try:
        driver.implicitly_wait(1)
        exist_value = operateValue.split("|")[0]
        not_exist_value = operateValue.split("|")[1]
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None:
            return exist_value
    except TimeoutException as e:
        return not_exist_value
    except Exception as e:
        raise e

def ifExistThenChooseOperateValue_diffPosition(locationType, locatorExpression, operateValue, *arg):
    # 返回值格，需填写两个位置信息，中间以"[]"分隔。两个返回值择一，填入不同格中。
    # 表格操作值填写格式：元素存在时返回值|元素不存在时返回值
    # 若元素存在，则返回表格操作值中，"|"之前的值，写入"[]"之前的坐标中；
    # 元素不存在，则返回"|"之后的值，写入"[]"之后的坐标中；
    global driver
    from selenium.common.exceptions import TimeoutException
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
    except TimeoutException as e:
        return not_exist_return_value
    except Exception as e:
        raise e

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

def ifExistThenReturnOperateValue(locationType, locatorExpression, operateValue, *arg):
    # 若元素存在，则返回表格操作值
    # 若元素不存在，则清空表格中的值
    global driver
    from selenium.common.exceptions import TimeoutException
    try:
        if operateValue is None:
            return_operateValue = None
        else:
            return_num = operateValue.count("[]")
            return_operateValue = "[]"*return_num
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by=locationType, value=locatorExpression))
        if element is not None:
            return operateValue
    except TimeoutException as e:
        return return_operateValue
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
        randContractNum = myInfo + randomNum(10)
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

def checkStateOfContract(operateValue):
    '''判断合同提交完成后生效状态'''
    global driver
    try:
        state = operateValue.split("|")[0]
        contractType = operateValue.split("|")[1]
        op1 = operateValue.split("|")[2]
        op2 = operateValue.split("|")[3]
        if state == "是" \
            and (contractType == "备货采购类型" or contractType == "样机采购类型" or contractType == "项目采购类型"):
            return op2
        else:
            return op1
    except Exception as e:
        raise e

# ****************************************项目关键字：财务管理模块****************************************
def getHidenInfo(myText):
    '''一层层点击开下拉框折叠项
    :param myText: 标准格式为“x1|x2|x3”,依次点击“x1”、“x3”、“x3”
    '''
    myTextArray = myText.split("|")
    for i in range(len(myTextArray)):
        sleep(0.5)
        # myXpath = "//*[contains(text(),\'" + myTextArray[i] + "\')]"
        myXpath = "//span[contains(text(),\'" + myTextArray[i] + "\')]"
        highlightElement("xpath", myXpath)
        click_Obj("xpath",myXpath)
        sleep(0.5)

def getInfoWanted(locationType,locatorExpression,myTest):
    '''遍历表格数据，校验详情页面是否存在指定字段（通常校验“备注”字段），暂不支持翻页
    :param locationType: 定位方式
    :param locatorExpression: 定位值
    :param myTest: 目标数据详情页面存在的唯一字段
    '''
    global driver
    try:
        elements = findElesbyMethod(driver,locationType,locatorExpression)
        print("********** 共有 ",len(elements)," 个待遍历元素 **********")
        loopTime = 1
        for i in elements:
            highlight(driver,i)
            i.click()
            switch_to_now_window(1)
            loadPage()

            try:
                myXpath = "//span[.='" + myTest + "']"
                waitVisibilityOfElementLocated("xpath",myXpath)
                sleep(0.5)
                print("********** 已找到目标记录！ **********")
                return
            except:
                close_page()
                switch_to_now_window(0)

            try:
                assert loopTime != len(elements), u"未找到目标记录！"
                loopTime += 1
            except AssertionError as e:
                print("********** 未找到目标记录！ **********")
                raise AssertionError(e)

    except Exception as e:
        raise e

def getNextUser(attributeType,locatorExpression,userBefore):
    '''获取下一岗审批人，并依据模块循环规则写入“数据表”sheet相应字段
    :param attributeType:待获取元素的属性种类
    :param locatorExpression:待获取元素的xpath
    :param userBefore:已有的全部审批人信息
    :return:含下一岗的全部审批人信息（若已结束审批，则返回值格式为“xx*xx*..xx*”）
    '''
    strTransformed = ifExistThenReturnAttribute_pinyin("xpath",locatorExpression,attributeType)

    if userBefore == "循环初始化":
        myUser = strTransformed
    else:
        myUser = userBefore + "*" + strTransformed

    return myUser

def setAmountOfPayment(myInfo):
    '''付款申请新建中，关联完采购订单后修改申请金额
    :param myTitle: 采购订单号
    :param amount: 申请金额
    '''
    myTitle = myInfo.split("|")[0]
    amount = myInfo.split("|")[1]
    myXpath = '//td[@title="' + myTitle +'"]/following-sibling::td[last()]/input'
    sendkeys_To_Obj("xpath",myXpath,amount)
    pageKeySimulate("xpath",myXpath,"page_tab")

def ifExist_pageKeySimulate(locationType,locatorExpression,keyType):
    ''' 若元素存在，则进行页面滑动 '''
    try:
        driver.implicitly_wait(1)
        element = WebDriverWait(driver, 1).until(lambda x: x.find_element(by = locationType, value = locatorExpression))
        pageKeySimulate(locationType,locatorExpression,keyType)
    except Exception as e:
        pass


# ****************************************项目关键字：进出口合同模块****************************************
def getNumWanted(locationType, locatorExpression, myNum):
    '''遍历数据，查找页面是否存在指定合同号，支持翻页
    :param locationType: 定位方式
    :param locatorExpression: 定位值
    :param myNum: 目标数据中唯一字段的text属性
    '''
    global driver
    from selenium.common.exceptions import TimeoutException
    try:
        myXpath = "//a[]"
        xpath_combination_click("text()", myXpath, myNum)

    except TimeoutException as e:
        # 若没有找到指定元素，开始翻页操作
        try:
            elements = findElesbyMethod(driver,locationType,locatorExpression)
            print("********** 共有 ",int(elements[-1].text)-1," 个待遍历元素 **********")
            loopTime = 1
            for i in range(int(elements[-1].text)-1):
                highlightElement("xpath",'//*[contains(@class,"pagination")]/span[not(text()="共0条记录")]/../ul/li[last()][not(@class="disabled")]')
                click_Obj("xpath",'//*[contains(@class,"pagination")]/span[not(text()="共0条记录")]/../ul/li[last()][not(@class="disabled")]')
                loopTime += 1
                loadPage()

                try:
                    xpath_combination_click("text()", myXpath, myNum)
                    print("********** 已找到目标记录！ **********")
                    return
                except TimeoutException as e:
                    assert loopTime < int(elements[-1].text), u"未找到指定的单据号！"
                except Exception as e:
                    raise e
        except Exception as e:
            raise e
    except Exception as e:
        raise e

# ****************************************项目关键字：个人中心****************************************
def setCheckBox(myText):
    '''点击多个check box
    :param myText: 标准格式为“x1|x2|x3”,依次点击“x1”、“x3”、“x3”
    '''
    myTextArray = myText.split("|")
    try:
        for i in range(len(myTextArray)):
            myXpath_A = "//div[.='" + myTextArray[i] + "']/preceding-sibling::input[@class='check']"
            myXpath_B = "//span[.='" + myTextArray[i] + "']/input"

            errInfo = "未找到的元素：%s！" %myTextArray[i]
            try:
                element = WebDriverWait(driver, 1).until(
                    lambda x: x.find_element(by="xpath", value=myXpath_A), errInfo)
                assert element.is_enabled() == True
            except Exception as e:
                element = WebDriverWait(driver, 1).until(
                    lambda x: x.find_element(by="xpath", value=myXpath_B), errInfo)
            # highlight(driver, element)
            element.click()
    except Exception as e:
        raise e


# **************************************** 日常办公 ****************************************
def getInfoNeeded(myText):
    ''' 针对待办记录中无法用单据号查询情况，遍历当前页获取对应单据信息
    :param myText: 单据号
    '''
    try:
        global driver
        xp = "(//th[.=\'签报单号\']/../../../tbody//tr)"
        trs = findElesbyMethod(driver, "xpath", xp)
        for i in range(len(trs)):
            xps = xp + "[" + str(i+1) + "]//a"
            el = findElebyMethod(driver, "xpath", xps)
            if el.text == myText:
                el.click()
                return
            assert i + 1 < len(trs), "当前页未找到单据记录！"
    except AssertionError as e:
        raise e
    except Exception as e:
        raise e


# **************************************** 组合功能 ****************************************
def getApprovalFlow(returnFlag=None):
    ''' （单据详情页）获取审批流详情，避免每岗审批后均要查询
    :param flag: 自动审批跳出函数标志位，防止自动审批情况下查不到审批流信息报错
    :return: 审批流信息
    '''
    # 当用户姓名中有多音字或其他原因，会导致转换的拼音字符串与用户账号不对应。
    # 针对于此情况，函数会先从special_list字典中，检索用户对应的拼音字符串。若检索不到，则进行拼音转换。
    special_list = {
        u'李浩': 'lihao01',
        u'曾蓉琴': 'zengrongqin',
        u'朴贤国': 'piaoxianguo',
        u'王洋': 'wangyang01',
        u'张鑫': 'zhangxin01',
        u'曾春娇': 'zengchunjiao'
    }
    trans = False
    try:
        waitVisibilityOfElementLocated("xpath", "//span[contains(text(), '审批状态')]")
    except Exception as e:
        if returnFlag is not None:
            raise e
        else:
            return

    try:
        scroll_slide_field("xpath", "//span[contains(text(), '审批状态')]")

        finalStr = ""
        eles = findElesbyMethod(driver, "xpath", "//td[.='待处理']")

        for i in range(len(eles)):
            xp = "((//td[.='待处理'])[%d]/../td)[2]" %(i + 1)
            element = findElebyMethod(driver, "xpath", xp)

            ip = "((//td[.='待处理'])[%d]/../td)[2]//i" %(i + 1)
            # c = getAttribute("xpath", ip, "class")
            c = findElesbyMethod(driver, "xpath", ip)[0].get_attribute("class")

            if c != "hideCss":
                for (k, v) in special_list.items():
                    if element.text == k:
                        finalStr += v + "*"
                        trans = True
                if trans == False:
                    finalStr += pinyinTransform(element.text) + "*"
                trans = False

        print("finalStr: ", finalStr)
        return finalStr
    except Exception as e:
        raise e


def loginProcess(username, password, loginTime=15):
    ''' 登陆流程 '''
    try:
        sendkeys_To_Obj("xpath", '//input[@ng-model="user_name"]', username)
        sendkeys_To_Obj("xpath", '//input[@ng-model="password"]', password)
        click_Obj("xpath", '//*[contains(@value,"登录") or .="登录"]')
        sleep(1)
        wait_elements_vanish("xpath", '//*[contains(@value,"登录") or .="登录"]', loginTime)
        wait_elements_vanish("xpath", '//td[.="没有单据内容"]', 30)
        # sleep(1)
        # ifExistThenClick("xpath", '//span[.="我的单据"]')
        # ifExistThenClick("xpath", '//span[.="我的单据(新)"]')
        loadPage()
        waitVisibilityOfElementLocated("xpath", '//div[.="我的单据"]')
    except Exception as e:
        raise e


def checkToLogin(userInfo):
    ''' 含判断机制的登陆
    :param userInfo: 用户名|密码
    '''
    userInfo = userInfo.split("|")
    username = userInfo[0]
    password = userInfo[1]
    try:
        # el存在，则已登录
        el = findElebyMethod(driver, "xpath", '//li[@class="nav-item dropdown"]/a/span/span', timeout=1)
        assert username in el.text
        # username即当前用户，跳出
        return
    except AssertionError:
        # username非当前用户，须先登出再重新登陆
        click_Obj("xpath", '//li[@class="nav-item dropdown"]')
        click_Obj("xpath", '//a[.="退出"]')
        loginProcess(username, password)
    except Exception:
        # el不存在
        try:
            loginEl = findElebyMethod(driver, "xpath", '//input[@ng-model="user_name"]', timeout=1)
            # 登陆界面，可直接登陆
        except:
            # 须从调起driver开始
            open_browser("chrome")
            maximize_browser()
            # visit_url(userInfo[2])
        finally:
            loginProcess(username, password)


# [SAP]
# **************************************** 基础配置 ****************************************
@SAPException
def createObject(info):
    ''' 调起SAP服务，20190521
    :param path: SAP执行文件路径
    :param env: 本地登陆环境名
    :return: 全局变量session
    '''
    import subprocess
    import win32com.client
    global session, MS, connection
    info = info.split("|")
    path = info[0]
    env = info[1]

    subprocess.Popen(path)
    time.sleep(1)

    # 最长等待时间（须≥2）
    MS = 10

    loopTime = 3
    for i in range(loopTime):
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            print("********** SAP成功启动！ **********")
            assert type(SapGuiAuto) == win32com.client.CDispatch
            break
        except AssertionError as e:
            return
        except Exception as e:
            if i + 1 < loopTime:
                continue
            else:
                raise e

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    # help(application.OpenConnection)
    # Function openConnection(descriptionString As String, sync As Boolean False, raiseAsBoolean = True)

    # connection = application.OpenConnection(env, True)
    #
    # if not type(connection) == win32com.client.CDispatch:
    #     application = None
    #     SapGuiAuto = None
    #     return

    for i in range(loopTime):
        try:
            connection = application.OpenConnection(env, True)
            assert type(connection) == win32com.client.CDispatch
            break
        except AssertionError as e:
            application = None
            SapGuiAuto = None
            return
        except Exception as e:
            if i + 1 < loopTime:
                continue
            else:
                raise e


    for i in range(loopTime):
        try:
            session = connection.Children(0)
            assert type(session) == win32com.client.CDispatch
            break
        except AssertionError as e:
            connection = None
            application = None
            SapGuiAuto = None
            return
        except Exception as e:
            if i + 1 < loopTime:
                continue
            else:
                raise e

@SAPException
def saplogin(info):
    ''' 登陆SAP，20190521 '''
    global session, AW
    info = info.split("|")
    userName = info[0]
    passWord = info[1]

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = userName
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = passWord
    session.findById("wnd[0]").sendVKey(0)

    try:
        # 等待“多次登陆”弹出框2s
        waitUntil(session.Children.count==2, "False", maxSec="2|pass")
        # 父对象，'/app/con[i]'
        ''' 方法一 .Parent '''
        # p = self.session.Parent
        # try:
        #     if p.findById("ses[0]/wnd[1]"):
        #         self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        #         self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        #         self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        # except:
        #     pass

        ''' 方法二 .ActiveWindow.FindAllByName '''
        updateActiveWindow()
        if AW.FindAllByName("wnd[1]", "GuiModalWindow").count == 1:
            AW.findById("usr/radMULTI_LOGON_OPT2").select()
            AW.findById("usr/radMULTI_LOGON_OPT2").setFocus()
            AW.findById("tbar[0]/btn[0]").press()
            # session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            # session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    # 等待“信息”弹出框1s
    try:
        waitObj("name|wnd[1]|GuiModalWindow", "pass", "1|信息")
        btn = session.findById("wnd[1]/tbar[0]/btn[0]")
        btn.press()
    except:
        pass

    waitObj("name|wnd[0]|GuiMainWindow", "err", "SAP 轻松访问 中建信息")
    print("********** SAP成功登陆！ **********")

@SAPException
def updateActiveWindow():
    ''' 更新ActiveWindow，20190521 '''
    global AW, session
    AW = session.ActiveWindow
    return AW

def closeSAP():
    ''' 结束SAP（所有）进程 '''
    import psutil, os
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        # print('pid-%s,pname-%s' % (pid, p.name()))
        if p.name() == 'saplogon.exe':
            cmd = 'taskkill /F /IM saplogon.exe'
            # 无需结束sap进程时，注释下行
            os.system(cmd)

@SAPException
def createNewSession():
    ''' 创建新session（窗口） '''
    global session, connection
    oc = connection.children.count
    session.createSession()         # session数+1
    sleep(1)
    session = connection.children(oc)
    for i in range(2):
        updateActiveWindow()
        try:
            waitObj("name|titl|GuiTitlebar", "pass", "SAP 轻松访问 中建信息|3")
            return
        except:
            pass
    updateActiveWindow()
    waitObj("name|titl|GuiTitlebar", "err", "SAP 轻松访问 中建信息|3")

def closeAllSession():
    global connection, session
    if 'connection' not in locals().keys():
        return

    sessions = connection.children
    for i in range(sessions.count):
        # 遍历 sessions.count 个session
        for j in range(10):
            try:
                wndId = "wnd[" + str(j) + "]"
                window = sessions[i].findById(wndId)
                window.close()
                try:
                    # 存在“注销”弹出框，则为最后一个window
                    msgBox =  sessions[i].findById("wnd[" + str(j+1) + "]")
                    assert msgBox.text == "注销"
                    msgBox.findById("usr/btnSPOP-OPTION1").press()
                    session = None
                    return
                except:
                    pass
            except:
                break

# **************************************** 基本操作 ****************************************
@SAPException
def performObj(performType, info, text=None):
    ''' 操作对象
    :param performType: 操作类型（输入/左击/勾选框/聚焦）
    :param info: 属性类型（id/name）|属性值
    :param text: 输入内容（performType为“输入”时）/对象text属性值（performType非“输入”时），选填
    '''
    global session, AW
    myInfo = info.split("|")

    if len(myInfo) == 4:
        # info: "name"|name|type|text
        obj = getObjsByNameAndText(myInfo[1] + "|" + myInfo[2], myInfo[3])
    elif performType in ["左击", "聚焦", "选择"] and text:
        obj = getObjsByNameAndText(info.split("|", 1)[1], text)
    else:
        obj = getObj(myInfo[0], info.split("|", 1)[1])

    if performType == "输入":
        obj.text = text
    elif performType == "下拉框":
        obj.key = text
    elif performType == "左击":
        obj.press()
    elif performType == "双击":
        obj.doubleClick()
    elif performType == "勾选框":
        # text = -1（勾选）/0（取消勾选）
        if obj.type in ["GuiCheckBox"]:
            # 示例type：GuiCheckBox
            obj.selected = int(text)
        elif obj.type in ["GuiShell"]:
            # 示例type：GuiShell
            obj.modifyCheckbox(0, "SEL", int(text))
            obj.triggerModified()
    elif performType == "聚焦":
        obj.setFocus()
    elif performType == "选择":
        obj.Select()
    elif performType == "模拟键盘":
        if text in keySimulated.keys():
            obj.sendVKey(keySimulated[text])
        else:
            obj.sendVKey(text)
    elif performType == "关闭":
        obj.close()

@SAPException
def performOnTable(performType, objInfo, inputInfo):
    ''' 根据同一类对象（同列）的不同index（不同行），操作table表格
    :param performType: 操作类型
    :param objInfo: name|type（表格类型为“GuiShell”时，后面还有“列信息”）
    :param inputInfo: text1%n1|text2%n2|...
                        (text为输入值；“%n”为选填项，对应表格index，从0开始)
    '''
    global AW
    II = inputInfo.split("|")
    OI = objInfo.split("|")
    # 默认从第一行开始
    initInfo = II[0].split("%")
    rowCount = 0 if len(initInfo) == 1 else int(initInfo[1])

    for i in range(len(II)):
        info = II[i].split("%")
        text = info[0]
        if i > 0:
            rowCount = (rowCount + 1) if len(info) == 1 else int(info[1])
        objs = AW.FindAllByName(OI[0], OI[1])

        if objs.count > 0 and objs[0].type == "GuiShell":
            # GuiShell类型表格
            obj = objs[0]
            if performType == "输入":
                obj.modifyCell(rowCount, OI[2], text)
        else:
            # 常规表格，通过index定位
            obj = objs[rowCount]
            if performType == "输入":
                obj.text = text
            elif performType == "下拉框":
                obj.key = text
            elif performType == "左击":
                obj.press()
            elif performType == "双击":
                obj.doubleClick()
            elif performType == "勾选框_A":
                # text = -1（勾选）/0（取消勾选）
                # 示例type：GuiCheckBox
                obj.selected = int(text)
            elif performType == "聚焦":
                obj.setFocus()
            elif performType == "选择":
                obj.Select()
            elif performType == "模拟键盘":
                obj.sendVKey(keySimulated[text])

@SAPException
def performToolbar(performType, info):
    ''' 操作type为GuiShell对象
    :param performType: 操作类型（提交/全部选择/清空选择/查看详情）
    :param info: 属性类型（id/name）|属性值
    '''
    global session, AW
    myInfo = info.split("|")

    if len(myInfo) == 4:
        # info: "name"|name|type|text
        obj = getObjsByNameAndText(myInfo[1] + "|" + myInfo[2], myInfo[3])
    else:
        obj = getObj(myInfo[0], info.split("|", 1)[1])

    if performType == "提交":
        obj.pressToolbarButton("ZPOST")
    elif performType == "全部选择":
        obj.pressToolbarButton("ZSALL")
    elif performType == "清空选择":
        obj.pressToolbarButton("ZDSAL")
    elif performType == "查看详情":
        obj.selectContextMenuItem("&DETAIL")

@SAPException
def getObj(case, info):
    ''' 获取对象
    :param case: 属性
    :param info: 属性值
    '''
    global AW
    info = info.split("|")
    typeInfo = case + "|" + info[0]
    typeInfo += ("|" + info[1]) if len(info) > 1 else ""
    waitObj(typeInfo, "err")

    if case == "id":
        obj = AW.findById(info[0])
        return obj
    elif case == "name":
        # obj = AW.FindByName(info[0], info[1])
        obj = AW.FindAllByName(info[0], info[1])[0]
        return obj

@SAPException
def getObjsByNameAndText(info, text):
    ''' 通过FindAllByName获取对象集合，再由text唯一定位
    :param info: 属性值（name|type）
    :param text: text属性值
    '''
    global AW
    typeInfo = "name|" + info
    waitObj(typeInfo, "err")

    info = info.split("|")
    objs = AW.FindAllByName(info[0], info[1])
    for i in range(objs.count):
        if objs[i].text == text:
            return objs[i]
        assert i + 1 == objs.count, \
            "该页面未找到name为“%s”、type为“%s”、text为“%s”的对象！" %(info[0], info[1], text)

@SAPException
def getText(case, info):
    ''' 获取对象text '''
    obj = getObj(case, info)
    return obj.text

@SAPException
def getNumInText(numInfo):
    import re
    num = re.findall('\d+', numInfo)
    return num[0] if len(num) == 1 else num[1]

# **************************************** 等待和校验 ****************************************
@SAPException
def waitObj(typeInfo, valueReturned, extraInfo=""):
    ''' 查找id，循环等待对象，20190522
    :param typeInfo: 属性（id/name）|属性值
    :param valueReturned: 未等到结果时结束方式（err：报错/pass：通过）
    :param extraInfo（选填）:可填写最长等待时间（默认为MS），及对象text属性（有值校验，不填不检验），顺序不限，用“|”隔开
    '''
    import re
    global AW, MS
    # 数据初始化
    updateActiveWindow()
    typeInfo = typeInfo.split("|")

    num = re.findall('\d+', extraInfo)
    num = num[0] if len(num) else ""
    maxSec = MS if not num else int(num)
    text = extraInfo.replace(num, "").replace("|", "")

    for i in range(maxSec):
        try:
            if typeInfo[0] == "id":
                obj = AW.findById(typeInfo[1])
            elif typeInfo[0] == "name":
                obj = AW.FindByName(typeInfo[1], typeInfo[2])

            # 校验text属性
            if text:
                assert obj.text == text

            return
        except:
            try:
                assert maxSec != i + 1, \
                    str(maxSec) + " s内未找到 '" + typeInfo[0] + "' 为 '" + typeInfo[1] + "' 的对象！"
                time.sleep(1)
            except AssertionError as e:
                if valueReturned == "err":
                    raise AssertionError(e)
                elif valueReturned == "pass":
                    pass
            except Exception as e:
                raise e

@SAPException
def waitUntil(event, condition, maxSec):
    ''' 循环等待（event是否发生与condition一致时，等待，否则跳出），20190521
    :param event: 待判断主体（布尔类型）
    :param condition: True/False
    :param maxSec: 最长循环等待时间（≥2）|未等到结果（err：报错/pass：通过）
    '''
    ms = maxSec.split("|")
    maxSec = int(ms[0])
    result = ms[1]
    for i in range(maxSec):
        try:
            assert maxSec != i + 1, \
                str(maxSec) + " s内均未等到指定事件发生！"
            if bool(event) == eval(condition):
                time.sleep(1)
            else:
                return
        except AssertionError as e:
            if result == "err":
                raise AssertionError(e)
            elif result == "pass":
                pass
        except Exception as e:
            raise e

@SAPException
def checkText(textExpected, realText):
    ''' 校验
    :param textExpected: 预期值
    :param realText: 实际值
    '''
    assert textExpected == realText, \
        "预期值 " + textExpected + " 与实际值 " + realText + " 不符！"

# **************************************** 项目关键字 ****************************************
# 采购申请
@SAPException
def chooseToReturn(infoA, infoB, infoC):
    infoA = infoA.split("|")
    infoB = infoB.split("|")
    for i in range(len(infoB)):
        if "|" in infoC:
            for str in infoC.split("|"):
                if str == infoB[i]:
                    return infoA[0]
        else:
            if infoC == infoB[i]:
                return infoA[0]
    VR = infoA[1] if len(infoA) > 1 else None
    # if len(infoA) > 1:
    #     VR = infoA[1]
    return VR

@SAPException
def chooseHowToTrans(var):
    var = var.split("|", 1)
    account = 0
    account += 1 if var[0] == "0" else 0
    for str in var[1].split("|"):
        if str in ["200", "300"]:
            account += 1
            break
    VR = "" if account == 2 else "zif001"
    return VR
