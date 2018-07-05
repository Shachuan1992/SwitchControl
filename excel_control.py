import xlwt
import xlrd
import pandas as pd
import re
import telnetlib
import time
from commands import *

def read_excel(file,row,clo):
    wb = xlrd.open_workbook(filename=file)
    sheet = wb.sheet_by_name('光衰清单')
    return sheet.cell_value(row,clo)

def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.colour_index = 4
    font.height = height
    style.font = font
    return style

def write_excel(file,row,clo,data):
    f = xlwt.Workbook()
    shee1 = f.add_sheet('Test',cell_overwrite_ok=True)
    shee1.write(row,clo,data,set_style('Times New Roman',220,True))
    f.save(file)


def Telnet_Switch():
    df = pd.read_excel('test.xls')
    account = input('请输入要查询的宽带账号\n')
    olt_ip = str(re.findall(r'\d+.\d+.\d+.\d+', str(df[df['账号'] == account]['IP'])))
    device = str(re.findall(r'[A-Z][A-Z0-9]{3,7}', str(df[df['账号'] == account]['设备型号'])))
    pon = str(re.findall(r'\d/\d/\d', str(df[df['账号'] == account]['PON口'])))
    IP = eval(olt_ip.strip('[[]]'))
    DEVICE_ID = eval(device.strip('[[]]'))
    PON_PORT = eval(pon.strip('[[]]'))
    print(IP,DEVICE_ID,PON_PORT)
    command_list = ['conf','terminal','exit']
    tn = telnetlib.Telnet(host=IP,port=2001)
    for command in command_list:
        write_command(tn,command)
        time.sleep(2)
    result = tn.read_very_eager().decode()