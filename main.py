# -*- coding:utf-8 -*-
"""
这是自动查光衰的程序

最终要实现的目标：
    通过输入宽带账号，可以自动调用SecureCRT执行光衰查询命令，可以自动判断设备型号（华为 or 中兴）并执行。
    并可以在excel表中自动标注符合标准的单元格

"""
__author__ = 'shachuan'


import telnetlib
import time
from commands import *
from excel_control import *


port_office = 8000
port_home = 2001


Telnet_Switch()
