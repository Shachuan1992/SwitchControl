# -*- coding:utf-8 -*-

__author__ = 'shachuan'


import telnetlib
import time
from commands import *

host = '127.0.0.1'
port = 5000


tn = telnetlib.Telnet(host,port=port)
write_command(tn,'conf')
time.sleep(1)
write_command(tn,'terminal')
time.sleep(1)
write_command(tn,'exit')
result = tn.read_very_eager().decode()
