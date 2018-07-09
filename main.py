# -*- coding:utf-8 -*-
"""
光衰自动查询程序

最终要实现的目标：
    通过输入宽带账号，可以自动调用Telnet协议，远程登录OLT执行光衰查询命令，可以自动判断设备型号（华为 or 中兴）并执行不同的命令。
    并可以在excel表中自动标注符合标准的单元格

"""
__author__ = 'shachuan'

from telnet import *

file = '2018.7.5光衰报表.xlsx'

state = True
while state == True:
    account = input("请输入要查询的宽带账号，按q退出\n")
    if account == ('q'):
        print('谢谢使用')
        break
    else:
        get_info(file, account)
        telnet()