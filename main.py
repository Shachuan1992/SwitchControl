# -*- coding:utf-8 -*-
"""
这是自动查光衰的程序，可以一键登录olt。
内部集成了常用命令集合，可以自动反馈光衰
"""
__author__ = 'shachuan'


import telnetlib
import time
from commands import *

host = '127.0.0.1'
port = 5000

command_list = ['conf','terminal','exit']
tn = telnetlib.Telnet(host,port=port)

for command in command_list:
    write_command(tn,command)
    time.sleep(2)

result = tn.read_very_eager().decode()
