#encoding = utf - 8
# 用于存放等待页面元素方法

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep

class WaitUtil(object):
    # 映射定位方式字典对象
    def __init__(self,driver):
        self.locationTyoeDict = {
            "xpath":By.XPATH,
            "id":By.ID,
            "name":By.NAME,
            "css_selector":By.CSS_SELECTOR,
            "class_name":By.CLASS_NAME,
            "tag_name":By.TAG_NAME,
            "link_text":By.LINK_TEXT,
            "partial_link_text":By.PARTIAL_LINK_TEXT
        }
        # 初始化driver
        self.driver = driver
        # 创建显式等待实例对象,等待时间待全局化处理，TODO
        self.wait = WebDriverWait(self.driver,15)

    def presenceOfElementLocated(self,locatorMethod,locatorExpression,*args):
        '''
        显式等待页面元素出现在DOM中，但不一定可见，存在则返回元素对象
        :param locatorMethod: 定位方法
        :param locatorExpression: 定位表达式
        :param args:
        :return: 页面元素对象
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locatorMethod, locatorExpression)
        try:
            if locatorMethod.lower() in self.locationTyoeDict:
                element = self.wait.until(
                    EC.presence_of_element_located((self.locationTyoeDict[locatorMethod.lower()],locatorExpression)), errInfo)
                return element
            else:
                raise TypeError(u'未找到定位方式，请确认定位方法使用是否正确')
        except Exception as e:
            raise e

    def visibilityOfElementLocated(self,locationType,locationExpression,*args):
        '''
        显式等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象
        :param locationType: 定位方法
        :param locationExpression: 定位表达式
        :param args:
        :return: None
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locationType, locationExpression)
        try:
            el = self.wait.until(EC.visibility_of_element_located((self.locationTyoeDict[locationType.lower()],
                                                                   locationExpression)), errInfo)
            print('********** 元素是否可操作：',el.is_enabled()," **********")
            sleep(0.5)
            # return el
        except Exception as e:
            raise e

    def frameToBeAvailableAndSwitchToIt(self,locationType,locationExpression,*args):
        '''
        检查frame是否存在，存在则切换到frame控件中
        :param locationType: 定位方法
        :param LocationExpression: 定位表达式
        :param args:
        :return: None
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locationType, locationExpression)
        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((self.locationTyoeDict[locationType.lower()],locationExpression)), errInfo)
        except Exception as e:
            # 抛出异常给上层调用者
            raise e

if __name__ == "__main__":
    from selenium import webdriver
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get("http://kintergration.chinacloudapp.cn:9002/login.html")

    waitUtil = WaitUtil(driver)
    waitUtil.presenceOfElementLocated("name","user_name")
    waitUtil.visibilityOfElementLocated("name","user_name")

    print('**********等待页面元素成功**********')

    driver.quit()