# -*- coding: utf-8 -*-
import time
import datetime
import ddddocr
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains

service = ChromeService(executable_path=r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')  # chrome驱动路径
chrome = webdriver.Chrome(service=service)

# 配置项

name = '123456'  # 账号
pwd = '132456l'  # 密码
safe_time = 3  # 安全间隔时间
servernum = 0  # 选择服务器 填： 0,1,2,3
retry = True  # 失败重试
mode = 1  # 查询模式
servers = ['http://127.0.0.1/', 'http://127.0.0.1/', 'http://127.0.0.1/', 'http://127.0.0.1/']


def print_INFO(message):
    print('[' + datetime.datetime.now().strftime('%H:%M:%S') + ']' + message)


def print_ERROR(error):
    print('[' + datetime.datetime.now().strftime('%H:%M:%S') + ']' + "\033[1;31m" + error + " \033[0m")


def print_Exception(e):
    print("\033[1;31m异常!\033[0m\n")
    print(e)


def recognize():
    if chrome.find_element(By.ID, 'icode').screenshot('img.png'):  # 捕获验证码
        # 验证码识别
        with open('img.png', 'rb') as f:
            img = f.read()
        ocr = ddddocr.DdddOcr()
        result = ocr.classification(img)
        print('[' + datetime.datetime.now().strftime('%H:%M:%S') + ']' + "验证码" + result)
        return result


def auto_Login():
    chrome.get(servers[servernum])
    print_INFO("尝试登录" + str(servernum) + '号服务器')
    if chrome.title == 'ERROR - 出错啦！':
        return False
    try:
        chrome.find_element(By.ID, 'txtUserName').send_keys(name)
        chrome.find_element(By.ID, 'TextBox2').send_keys(pwd)
    except:
        print_Exception(Exception)
        return False
    chrome.find_element(By.ID, 'txtSecretCode').send_keys(recognize())
    try:
        chrome.find_element(By.ID, 'Button1').click()
    except:
        print_Exception(Exception)
        return False
    if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
        print_ERROR('跳转失败')
        return False
    return True


def get_score():
    try:
        hovertarget = chrome.find_element(By.XPATH, '/html/body/div/div[1]/ul/li[5]/a/span')
        ActionChains(chrome).move_to_element(hovertarget).perform()
        chrome.find_element(By.XPATH, '/html/body/div/div[1]/ul/li[5]/ul/li[4]/a').click()
    except:
        print_Exception(Exception)
        print_ERROR('成绩查询按钮点击失败')
        return False
    if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
        auto_Login()
        return False
    time.sleep(5)
    try:
        chrome.switch_to.frame('zhuti')
        chrome.find_element(By.ID, 'btn_zcj').click()
        if chrome.title == 'ERROR - 出错啦！' or chrome.title == '欢迎使用正方教务管理系统！请登录':
            chrome.switch_to.default_content()
            while 1:
                if auto_Login():
                    break
            return False
        # 保存到excel
        work_book = openpyxl.Workbook()
        shell = work_book.worksheets[0]
        trs = chrome.find_elements(By.XPATH, '/html/body/form/div[2]/div/span/div[1]/table[1]/tbody/tr')
        trnum = 1
        for tr in trs:
            tdnum = 1
            while 1:
                tdXPATH = '/html/body/form/div[2]/div/span/div[1]/table[1]/tbody/tr[' + str(trnum) + ']/td[' + str(tdnum) + ']'
                shell.cell(trnum, tdnum, chrome.find_element(By.XPATH, tdXPATH).text)
                tdnum += 1
                if tdnum == 20:
                    break
            trnum += 1
        work_book.save('score.xlsx')
        chrome.switch_to.default_content()
    except:
        print_Exception(Exception)
        print_ERROR('成绩获取错误')
        return False
    return True


def main():
    if mode == 1:
        print_INFO('开始查询成绩')
        chrome.maximize_window()
        while 1:
            trytimes = 0
            succeed = False
            while 1:
                trytimes += 1
                if auto_Login():
                    succeed = True
                    print_INFO('登录成功')
                    break
                else:
                    print_ERROR('尝试登录失败')
                    if retry:
                        if trytimes > 10 and succeed == False:
                            print('\033[0;32m已经为你尝试了' + str(
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
                trygetscore += 1
                status = get_score()
                if status:
                    print_INFO('查询成功')
                    break
                else:
                    print_ERROR('查询失败')
                    if retry:
                        time.sleep(safe_time)
                        continue
                    else:
                        break
            if get_scoreFaile is True:
                continue
            else:
                a = input()
    elif mode == 2:
        print_INFO('该功能正在开发')



main()
