# -*- coding:utf-8 -*-
"""
光衰自动查询程序

最终要实现的目标：
    通过输入宽带账号，可以自动调用Telnet协议，远程登录OLT执行光衰查询命令，可以自动判断设备型号（华为 or 中兴）并执行不同的命令。

    update:2018/7/9
            需要改进的地方：
                1、输入检测，检测输入的宽带账号是否正确
                2、登录检测，若是登录超时则返回错误原因
                3、最好是不用延时，以接收到的数据为执行命令的前提条件
                4、针对可能出现的错误情况，能保持程序运行不中断
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