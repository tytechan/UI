#encoding = utf - 8
# 用于存放剪切板操作方法

import win32clipboard as w        #win32clipboard无需通过pip安装，安装pywin32即可
import win32con

class Clipboard(object):
    '''
    模拟windows设置剪切板
    '''

    @staticmethod
    def getText():      #读取剪切板
        # 打开剪切板
        w.OpenClipboard()
        # 获取剪切板数据
        mydata = w.GetClipboardData(win32con.CF_TEXT)
        # 关闭剪切板
        w.CloseClipboard()

        return mydata

    @staticmethod
    def setText(mytring):       #设置剪切板内容
        # 打开剪切板
        w.OpenClipboard()
        # 清空剪切板
        w.EmptyClipboard()
        # 将数据写入剪切板
        w.SetClipboardData(win32con.CF_UNICODETEXT,mytring)
        # 关闭剪切板
        w.CloseClipboard()