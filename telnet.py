import xlwt
import xlrd
import pandas as pd
import re
import telnetlib
import time

def write_command(instance,command):
    finish = '\r'
    instance.write(command.encode('ascii')+finish.encode('ascii'))

#读Excel函数
def read_excel(file,row=0,clo=0):
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

#从光衰表中获取OLTP的IP，设备型号Device_ID，Pon口PON_PORT以及ONU_ID
def get_info(file,account):
    global IP,DEVICE_ID,ONU_ID,SLOT,BOARD,PORT#声明全局变量
    df = pd.read_excel(file)#打开光衰表文件
    olt_ip = str(re.findall(r'\d+.\d+.\d+.\d+', str(df[df['账号'] == account]['OLT-IP地址'])))#使用正则表达式匹配IP
    device = str(re.findall(r'[A-Z][A-Z0-9]{3,7}', str(df[df['账号'] == account]['设备型号'])))#使用正则表达式匹配设备型号
    pon = str(re.findall(r'\d/\d/\d', str(df[df['账号'] == account]['PON口'])))#使用正则表达式匹配Pon口
    onu = str(re.findall(r'\d{1,2}\.',str(df[df['账号'] == account]['ONUID'])))  # 使用正则表达式匹配Pon口
    IP = eval(olt_ip.strip('[[]]'))#去掉字符串两边的字符
    DEVICE_ID = eval(device.strip('[[]]'))#去掉字符串两边的字符
    PON_PORT = eval(pon.strip('[[]]'))#去掉字符串两边的字符
    ONU_ID = eval(onu.strip('[[]]')).rstrip('.')  # 去掉字符串两边的字符
    EPON = PON_PORT.split('/')
    SLOT = EPON[0]
    BOARD = EPON[1]
    PORT = EPON[2]


def telnet():

    tn = telnetlib.Telnet(host=IP, port=5000)   # 开启Telnet

    command_make(Device=DEVICE_ID)              # 根据设备型号，制作命令

    for command in commands:                    # 输入命令并执行
        write_command(tn,command)
        time.sleep(0.5)
        print(tn.read_very_eager())

#获取OLT返回的信息
def get_msg():
    pass

def command_make(Device):
    global commands
    commands = []
    if Device == 'MA5683T':
        command_interface = 'interface epon ' + SLOT + '/' + BOARD
        command_opticai_info = 'display ont optical-info ' + PORT + ' ' + ONU_ID
    elif Device == 'MA5680T':
        command_interface = 'interface epon ' + SLOT + '/' + BOARD
        command_opticai_info = 'display ont optical-info ' + PORT + ' ' + ONU_ID
    else:
        command_interface = 'conf'
        command_opticai_info = 'terminal'
        quit = 'do show running-config'
    commands.append(command_interface)
    commands.append(command_opticai_info)
    commands.append(quit)