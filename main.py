# -*- coding:utf-8 -*-
"""
这是自动查光衰的程序

最终要实现的目标：
    通过输入宽带账号，可以自动调用SecureCRT执行光衰查询命令，可以自动判断设备型号（华为 or 中兴）并执行。
    并可以在excel表中自动标注符合标准的单元格

"""
__author__ = 'shachuan'

from telnet import *

file = 'test.xls'
port_office = 5000
port_home = 2001

account = input('请输入要查询的宽带账号\n')

telnet_switch(file,account)
read_excel(file)
