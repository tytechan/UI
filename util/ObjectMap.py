#encoding = utf - 8
# 用于存放定位元素及操作的基本方法
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# import os


def findElebyMethod(driver, locateType, locatorExpression, timeout=10):
    errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locateType, locatorExpression)
    try:
        locateType = locateType.lower()
        if locateType == 'id':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.ID, locatorExpression)), errInfo)
        elif locateType == 'name':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.NAME, locatorExpression)), errInfo)
        elif locateType == 'classname':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.CLASS_NAME, locatorExpression)), errInfo)
        elif locateType == 'link_text':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.LINK_TEXT, locatorExpression)), errInfo)
        elif locateType == 'xpath':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.XPATH, locatorExpression)), errInfo)
        elif locateType == 'css_selector':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, locatorExpression)), errInfo)
        elif locateType == 'partial_link_text':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.PARTIAL_LINK_TEXT, locatorExpression)), errInfo)
        elif locateType == 'value':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.XPATH, "//*[contains(@value,'"+locatorExpression+"')]")), errInfo)
        elif locateType == 'text':
            specify_element = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.XPATH, "//*[text()='"+locatorExpression+"']")), errInfo)

        # myindex = int(index)
        # return specify_elements[myindex]
        return specify_element
    except Exception as e:
        raise e

def findElesbyMethod(driver, locateType, locatorExpression, timeout=10):
    errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locateType, locatorExpression)
    try:
        locateType = locateType.lower()
        if locateType == 'id':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.ID, locatorExpression)), errInfo)
        elif locateType == 'name':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.NAME, locatorExpression)), errInfo)
        elif locateType == 'classname':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, locatorExpression)), errInfo)
        elif locateType == 'link_text':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.LINK_TEXT, locatorExpression)), errInfo)
        elif locateType == 'xpath':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.XPATH, locatorExpression)), errInfo)
        elif locateType == 'css_selector':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, locatorExpression)), errInfo)
        elif locateType == 'partial_link_text':
            specify_elements = WebDriverWait(driver, timeout).until(
                EC.presence_of_all_elements_located((By.PARTIAL_LINK_TEXT, locatorExpression)), errInfo)

        # myindex = int(index)
        # return specify_elements[myindex]
        return specify_elements
    except Exception as e:
        raise e


def findEleByDetail(driver, locateType, locatorExpression, timeout=10):
    errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locateType, locatorExpression)
    try:
        element = WebDriverWait(driver, timeout).until(lambda x: x.find_element(by = locateType, value = locatorExpression), errInfo)
        return element
    except Exception as e:
        raise e

def findElesByDetail(driver, locateType, locatorExpression, timeout=10):
    errInfo = "未找到 '%s' 为 '%s' 的元素集合！" %(locateType, locatorExpression)
    try:
        elements = WebDriverWait(driver, timeout).until(lambda x: x.find_elements(by = locateType, value = locatorExpression), errInfo)
        return elements
    except Exception as e:
        raise e

def moveMouse(myX,myY):
    try:
        from pymouse import PyMouse
        m = PyMouse()
        m.position()
        m.move(myX, myY)
    except Exception as e:
        raise e

def highlight(driver,element):
    driver.execute_script("arguments[0].setAttribute('style',arguments[1]);",
                          element,"background:green;border:2px solid red;")

def addAttribute(driver,element,attributeName,value):
    driver.execute_script("arguments[0].%s = arguments[1]" %attributeName,element,value)

def setAttribute(driver,element,attributeName,value):
    driver.execute_script("arguments[0].setAttribute(arguments[1],arguments[2])",
                          element,attributeName,value)

def removeAttribute(driver,element,attributeName):
    driver.execute_script("arguments[0].removeAttribute(arguments[1])",
                          element,attributeName)


if __name__ == "__main__":
    from selenium import webdriver

    driver = webdriver.Chrome()
    print('打开浏览器')
    driver.maximize_window()

    driver.get('http://kintergration.chinacloudapp.cn:9002/login.html')
    print('打开网页')

    el = findEleByDetail(driver,"name","user_name")
    print(el.get_attribute("placeholder"))

    el1 = findEleByDetail(driver,"name","password")
    print(el1.get_attribute("placeholder"))

    driver.quit()