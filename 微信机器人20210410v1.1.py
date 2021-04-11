# 操作Web浏览器
from selenium import webdriver

# 操作桌面应用程序
import pywinauto
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse

# 操作时间
import time
import datetime

# 操作文件目录
import os

# 提醒消息模板
from string import Template

# 操作Excel
import numpy as np # 支持大量的维度数组与矩阵运算,此外也针对数组运算提供大量的数学函数库
import pandas as pd # 操作excel

# 操作进程
import psutil

# 下载腾讯文档表格
# url：腾讯文档的url
# downloadFloder：下载位置，Excel表格将导出到该文件夹
# fileName：腾讯文档的名称，也是Excel表格的名称

def downloadTencentExcel(url,downloadFloder,fileName):
    # selenium启动WebDriver
    driver = webdriver.Chrome(r"D:\00Data\00Code\20210408微信机器人_青年大学习\chromedriver_v89.0.4389.23.exe")
    # 注意chromedriver的版本与chrome的版本对应； chromedriver下载网址：http://chromedriver.chromium.org/downloads

    # 因为涉及多次导出，如果在“下载”文件夹中存在该文件，则先删除。
    # 否则，导出时，会有(1)、(2)等标记，文件名会发生改变
    if os.path.exists(os.path.join(downloadFloder, fileName)):
        os.remove(os.path.join(downloadFloder, fileName))

    # 设置默认等待时间为5s
    sleepTime=5
    
    # WebDriver访问腾讯文档url
    driver.get(url) 
    # 等待网页加载完成，等待5*5=25秒
    time.sleep(sleepTime*5)

    # 点击右上方“立即登录”按钮
    Button_Login = driver.find_element_by_id("header-login-btn")
    Button_Login.click()
    time.sleep(sleepTime)

    # 点击QQ头像，登录
    # 这里有个坑，定位不到iframe框架中的元素，需要先切换进iframe框架中，才能访问元素
    driver.switch_to.frame("login_frame")
    Button_QQPicture=driver.find_element_by_xpath('//a[@class="face"]')
    Button_QQPicture.click()
    time.sleep(sleepTime)

    # 点击导出按钮，导出Excel表格
    Button_Menu=driver.find_element_by_xpath('//div[@class="titlebar-icon titlebar-icon-more"]')
    Button_Menu.click()
    Button_export=driver.find_element_by_xpath("//span[contains(text(),'导出为')]")
    # Button_export.get_attribute('outerHTML')
    Button_export.click()
    time.sleep(sleepTime)

    Button_exportExcel=driver.find_element_by_xpath("//li[contains(text(),'本地Excel表格 (.xlsx)')]")
    Button_exportExcel.click()
    time.sleep(sleepTime*3)

    # 关闭浏览器
    driver.close()

# 确定哪些同学没有学习
# downloadFloder：存放Excel表格的文件夹
# fileName：Excel表格名称（包含后缀名）
# rows：需要忽略的行数，Excel表格前几行可能无关，需要剔除掉
# colName：列名称，该列判断是否学习
def NameList(downloadFloder,fileName,rows,colName):
    # 用于存放没有学习同学的数组
    nameList=[]
    # 读取Excel表格
    usersList=pd.read_excel(os.path.join(downloadFloder, fileName), sheet_name=0, header=0, index_col=None,na_values=['NA'])
    # 遍历每一行
    for i in usersList.index:
        # 第一行是无关项
        if i in range(rows):
            continue
        # 遍历每一行
        record=usersList.iloc[[i]]
        # 判断该列是否为空
        # if pd.isna(record["第十一季\n第四期"].values[0]):
        if pd.isna(record[colName].values[0]):
            # 如果为空，则追加到数组里
            nameList.append(record["姓名"].values[0])
    return nameList

# 获得进程ID
# name：进程的名称
# 利用进程的名称，获得该进程的ID
# 示例：GetWeChatPID('WeChat.exe')
def GetWeChatPID(name):
    # 获得全部进程的ID
    pids=psutil.pids()
    # 遍历全部进程ID
    for pid in pids:
        # 进程
        p = psutil.Process(pid)
        # 如果找到进程名称，则中止循环
        # print(p.name())
        if(p.name()==name):
            WeChatPID=pid
            break
    # 输出进程ID
    return WeChatPID

# 连接微信，返回微信左上角搜索框的位置
# PID：进程ID
def ConnectWechat(PID):
    # 连接微信，指定backend为uia模式
    app = Application(backend="uia").connect(process=PID)
    # 一定要加(backend="uia")
    # 获得微信窗口
    win = app[u'微信']
    # 将窗口最大化
    win.maximize()
    # 输出该窗口的全部控件的属性 
    # win.print_control_identifiers()
    # 定位到左上角搜索框，需要用到UISpy软件
    search = win.child_window(title=u"搜索", control_type="Edit")
    # 对搜索框进行描边
    search.draw_outline()
    # 获得搜索框的位置信息
    cords = search.rectangle()
    return cords

# 进入聊天界面
# cords：搜索框位置信息
# name：班级群名或者同学名称
def EnterChatInterface(cords,name):
    # 点击搜索框
    # 这里是个大坑，容易点击不上搜索框，这里我直接点击三次
    mouse.click(button='left', coords=(cords.left + 100, cords.top+10))
    mouse.click(button='left', coords=(cords.left + 100, cords.top+10))
    mouse.click(button='left', coords=(cords.left + 100, cords.top+10))
    # 输入名称
    send_keys(name)
    time.sleep(1)
    # 按Enter键，进入聊天界面
    send_keys('{ENTER}')

# 读取txt文件中的模板
def ReadTemplate(txtTemplateFile):
    # 读取txt模板
    with open(os.path.join(codeFolder,txtTemplateFile), 'r',encoding='utf-8') as f:
        txtTemplate=f.read()
    message=Template(txtTemplate)
    return message

# 获得当前时间，确定问候语是早上好、下午好、晚上好
def Greet():
    # 获得当前时间
    now = datetime.datetime.now()
    # 12点之前，为早上好
    if now.hour <= 12:
        return '早上好'
    # 12点之后，18点之前，为下午好
    elif now.hour <= 18:
        return '下午好'
    # 18点之后，为晚上好
    else:
        return '晚上好'

# 发送提醒消息，私发消息
# nameList：没有学习同学数组
def SendMessagePrivateChat(nameList):
    # 确定问候语
    greet=Greet()
    # 读取模板
    message=ReadTemplate(txtTemplateFile)
    for name in nameList:
        # 填写模板变量
        msg=message.substitute(Name=name[-2:len(name)],Greet=greet,DocumentName=documentName,DocumentUrl=url)
        # 输出消息
        print(msg)
        # 进入聊天姐买你
        EnterChatInterface(cords,name)
        time.sleep(1)
        # 键盘输入消息
        send_keys(msg)
        # ctrl+enter发送消息
        send_keys('^{ENTER}')

# 发送提醒消息，在群里发送消息
def SendMessageGroup(groupName,nameList):
    # 确定问候语
    greet=Greet()
    # 读取模板
    message=ReadTemplate(txtTemplateFile_qun)
    # 进入聊天界面
    EnterChatInterface(cords,groupName)
    
    # 输入提醒消息
    send_keys('自动提醒|青年大学习')
    send_keys('{ENTER}')
    # @没有学习同学
    for name in nameList:
        time.sleep(1)
        send_keys('@')
        time.sleep(1)
        # 输入同学姓名
        send_keys(name)
        time.sleep(0.3)
        # 按ENTER键，确定@该同学
        send_keys('{ENTER}')
        time.sleep(1)
    # 填写模板变量
    msg=message.substitute(Greet=greet,DocumentName=documentName,DocumentUrl=url)
    send_keys(msg)
    # ctrl+enter发送消息
    send_keys('^{ENTER}')

# 全局参数
# 文件名，Excel表格名称
documentName='青年大学习第十一季'
# 文件名，包含后缀名
fileName=documentName+'.xlsx'
# 班级群名
groupName='GIS班的小可爱们'
# 下载位置，Excel表格将导出到该文件夹
downloadFloder=r'C:\Users\vgp\Downloads'
# py代码所在的文件夹
codeFolder=r'D:\00Data\00Code\20210408微信机器人_青年大学习'
txtTemplateFile_qun='通知模板群发.txt'
txtTemplateFile='通知模板.txt'
emailFile='班级同学邮箱.xlsx'
url='https://docs.qq.com/sheet/DU2NzQ3Z0Qm5nVk5j'

def main():
    downloadTencentExcel(url,downloadFloder,fileName)
    nameList=NameList()
    print(nameList)
    # SendMessage(nameList)
    # 私发消息
    SendMessageQun(groupName,nameList)
    # 在群里发消息

if __name__ == '__main__':
    PID=GetWeChatPID()
    cords=ConnectWechat(PID)
    while True:
        h={10,15,17,21}
        now = datetime.datetime.now()
        # print(now.__format__('%H:%M:%S'))
        if now.hour in h and now.minute==0:
        # if True:
            main()
        time.sleep(50)


