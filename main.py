# -*- coding:utf-8 -*-
"""
这是自动查光衰的程序，可以一键登录olt。
内部集成了常用命令集合，可以自动反馈光衰
"""
__author__ = 'shachuan'


import telnetlib
import time
from commands import *
from excel_control import *

host = '127.0.0.1'
port_office = 8000
port_home = 2001

command_list = ['conf','terminal','exit']
tn = telnetlib.Telnet(host,port=port_home)

for command in command_list:
    write_command(tn,command)
    time.sleep(2)

result = tn.read_very_eager().decode()
