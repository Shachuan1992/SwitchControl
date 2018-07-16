# -*- coding:utf-8 -*-
"""
光衰自动查询程序

最终要实现的目标：
    通过输入宽带账号，可以自动调用Telnet协议，远程登录OLT执行光衰查询命令，可以自动判断设备型号（华为 or 中兴）并执行不同的命令。

    update:2018/7/9
            需要改进的地方：
                1、输入检测，检测输入的宽带账号是否正确
                2、登录检测，若是登录超时则返回错误原因
                3、针对可能出现的错误情况，能保持程序运行不中断。例如文件夹中没有光衰表，ONU不上线，IP地址无法登陆，没有此账号等
                4、根据查出来的光衰，能够在Excel里自动标注底色

    update:2018/7/10
                1、加入了文件判断函数，在执行查询操作前会查看当前目录是否存在xlsx后缀的文件，若存在的话会默认当成光衰报表并自动获取文件
                名。综合考虑没有使用if-else，而是使用了try-except语句，更专业。
    update:2018/7/16
                1、做了登录检测及文件检测，要是没找到正确的sheet名会返回错误，要是登录超时也会返回错误信息，但不会终止程序的执行
"""
__author__ = 'shachuan'

from telnet import *
from os import path

while True:
    try:
        file_path = path.dirname(__file__)
        isFileIn = scan_files(file_path, postfix='.xlsx')
        file_name = re.search(r".+\\(.+)", isFileIn[0]).group(1)
        account = input("请输入要查询的宽带账号，按q退出\n")
        if account == ('q'):
            print('谢谢使用')
            break
        else:
            try:
                get_info(file_name, account)
            except xlrd.biffh.XLRDError:
                print("错误！！！没有找到光衰清单的sheet，请打开光衰报表文件并修改sheet名为'光衰清单'\n")
                break
            except AttributeError:
                print("错误！！！宽带账号输入错误，请重新输入\n")
            else:
                telnet()
    except IndexError:
        print("错误！！！没有找到光衰文件！！！\n"
              "请确认光衰报表是否存在，并且目录下不要有多个excel文件,excel扩展名要用xlsx格式\n")
        break