# 正方教务爬虫
基于selenium的正方教务成绩爬虫

### 写在最前
***
这个爬虫是一个半成品，可以自动跳到成绩页面并保存成绩，后续功能略😁..
***

### 能力
😉自动登录✔
😍自动验证码识别填写✔
😊失败自动重试✔
😘自动保存成绩✔
自动评课×
自动选课×


# 使用
## step 1 下载
直接复制或下载本仓库里的spider.py文件到本地
## step 2 安装最新版Chrome浏览器（已有请跳过）
https://www.google.cn/intl/zh-CN/chrome/
## step 3 下载对应的驱动
下载地址  http://chromedriver.storage.googleapis.com/index.html
驱动安装教程 https://blog.csdn.net/m0_67575344/article/details/126142295

## step 4 安装所需模块
1. ddddocr
2. selenium

##### 命令行安装

```
pip install ddddocr
pip install selenium
```

##### 自动安装
使用pycharm自动安装


## step 5 运行
直接在编译器环境运行即可

## 已知bug
1.没有捕获验证码识别错误后抛出的异常
