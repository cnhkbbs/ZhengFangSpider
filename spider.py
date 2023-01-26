# -*- coding: utf-8 -*-
import time
import datetime
import ddddocr
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains

service = ChromeService(executable_path=r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')  # chrome驱动路径
chrome = webdriver.Chrome(service=service)

# 配置项

name = '123456'  # 账号
pwd = '132456l'  # 密码
realname = '123'  # 真实姓名
safe_time = 3  # 安全间隔时间
servernum = 0  # 选择服务器 填： 0,1,2,3
retry = True  # 失败重试

servers = ['http://127.0.0.1/', 'http://127.0.0.1/', 'http://127.0.0.1/', 'http://127.0.0.1/']
gnmkdms = ['', '', '', '']


def get_time():
    return '[' + datetime.datetime.now().strftime('%H:%M:%S') + ']'


def PrintERROR(e):
    print("\033[1;31m异常!\033[0m\n")
    print(e)


def recognize():
    if chrome.find_element(By.ID, 'icode').screenshot('img.png'):  # 捕获验证码
        # 验证码识别
        with open('img.png', 'rb') as f:
            img = f.read()
        ocr = ddddocr.DdddOcr()
        result = ocr.classification(img)
        print(get_time() + "验证码" + result)
        return result


def auto_Login():
    chrome.get(servers[servernum])
    print(get_time() + "尝试登录" + str(servernum) + '号服务器')
    if chrome.title == 'ERROR - 出错啦！':
        return False
    try:
        chrome.find_element(By.ID, 'txtUserName').send_keys(name)
        chrome.find_element(By.ID, 'TextBox2').send_keys(pwd)
    except:
        PrintERROR(Exception)
        return False
    chrome.find_element(By.ID, 'txtSecretCode').send_keys(recognize())
    time.sleep(1)
    try:
        chrome.find_element(By.ID, 'Button1').click()
    except:
        print(Exception)
        return False
    if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
        print(get_time() + "\033[1;31m跳转失败 \033[0m")
        return False
    return True


def get_score():
    try:
        # baseURL = chrome.current_url
        # scoreURL = baseURL[0:49] + 'xscjcx.aspx?xh=' + name + '&xm=' + realname + '&gnmkdm=' + gnmkdms[servernum]
        # print(scoreURL)
        # chrome.get(scoreURL)
        hovertarget = chrome.find_element(By.XPATH, '/html/body/div/div[1]/ul/li[5]/a/span')
        ActionChains(chrome).move_to_element(hovertarget).perform()
        chrome.find_element(By.XPATH, '/html/body/div/div[1]/ul/li[5]/ul/li[4]/a').click()
    except:
        print(Exception)
        print(get_time() + '\033[1;31m成绩查询按钮点击失败 \033[0m')
        return False
    if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
        auto_Login()
        return False
    time.sleep(5)
    try:
        chrome.switch_to.frame('zhuti')
        print(chrome.title)
        if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
            chrome.switch_to.default_content()
            while 1:
                if auto_Login():
                    break
            return False
        chrome.switch_to.default_content()

    except:
        print('\033[1;31m成绩获取错误\033[0m')
        return False
    return True


def main():
    chrome.maximize_window()
    while 1:
        trytimes = 0
        succeed = False
        while 1:
            trytimes += 1
            if auto_Login():
                succeed = True
                print(get_time() + "登陆成功\n")
                break
            else:
                print(get_time() + "\033[1;31m尝试登录失败\033[0m\n")
                if retry:
                    if trytimes > 10 and succeed == False:
                        print(get_time() + '\033[0;32m已经为你尝试了' + str(
                            trytimes) + '次登录, 全部登录失败。建议更换服务器或检查你的账号密码是否正确。\033[0m')
                    time.sleep(safe_time)
                    continue
                else:
                    break
        time.sleep(1)
        trygetscore = 0
        get_scoreFaile = False
        while 1:
            if trygetscore >= 5:
                get_scoreFaile = True
                break
            trygetscore +=1
            status = get_score()
            if status:
                print(get_time() + "查询成功")
                break
            else:
                print(get_time() + "\033[1;31m查询失败\033[0m")
                if retry:
                    time.sleep(safe_time)
                    continue
                else:
                    break
        if get_scoreFaile is True:
            continue
        else:
            a = input()




main()
