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
# def read_excel(file,row=0,clo=0):
#     wb = xlrd.open_workbook(filename=file)
#     sheet = wb.sheet_by_name('光衰清单')
#     return sheet.cell_value(row,clo)
#
# def set_style(name,height,bold=False):
#     style = xlwt.XFStyle()
#     font = xlwt.Font()
#     font.name = name
#     font.bold = bold
#     font.colour_index = 4
#     font.height = height
#     style.font = font
#     return style
#
# def write_excel(file,row,clo,data):
#     f = xlwt.Workbook()
#     shee1 = f.add_sheet('Test',cell_overwrite_ok=True)
#     shee1.write(row,clo,data,set_style('Times New Roman',220,True))
#     f.save(file)




#从光衰表中获取OLTP的IP，设备型号Device_ID，IP,PON_PORT以及ONU_ID
def get_info(file,account):
    global IP,DEVICE_ID,ONU_ID,PON,SLOT,BOARD,PORT#声明全局变量
    df = pd.read_excel(file,'光衰清单')#打开光衰表文件
    IP = str(re.search(r'(\d+.\d+.\d+.\d+)', str(df[df['账号'] == account]['OLT-IP地址']), re.M).group(1))              #使用正则表达式匹配IP
    PON = str(re.search(r'(\d/\d/\d)', str(df[df['账号'] == account]['PON口']), re.M).group(1))                         # 使用正则表达式匹配Pon口
    DEVICE_ID = str(re.search(r'([A-Z][A-Z0-9]{3,7})', str(df[df['账号'] == account]['设备型号']), re.M).group(1))       # 使用正则表达式匹配设备型号
    ONU_ID = str(re.search(r'^\d{1,4}\s+(\d{1,2})$', str(df[df['账号'] == account]['ONUID']), re.M).group(1))           # 使用正则表达式匹配ONUID
    PON_PORT = PON.split('/')
    SLOT = PON_PORT[0]
    BOARD = PON_PORT[1]
    PORT = PON_PORT[2]



def command_make(Device):
    global commands
    commands = []
    if (Device == 'MA5683T') or (Device == 'MA5680T'):
        commands.append('lyread')
        commands.append('read123')
        commands.append('enable')
        commands.append('config')
        commands.append('interface epon ' + SLOT + '/' + BOARD)
        commands.append('display ont optical-info ' + PORT + ' ' + ONU_ID)
    elif Device == 'C300':
        commands.append('show pon power onu-rx epon-onu_'+PON+':'+ONU_ID)
    else:
        print("错误！！！没有这种设备，需添加命令")



def telnet():

    tn = telnetlib.Telnet(host=IP, port=5000)   # 开启Telnet

    command_make(Device=DEVICE_ID)            # 根据设备型号，制作命令
    # 输入命令并执行
    for command in commands:
        write_command(tn,command)
        time.sleep(0.5)
    reply = (str(tn.read_very_eager(),encoding='utf-8'))
    if (DEVICE_ID == 'MA5683T') or (DEVICE_ID == 'MA5683T'):
        sample = reply.replace('(', '').replace(')', '')
        RxOptical = re.search(r"^\s+Rx optical powerdBm\s+:(.+)$", sample, re.M)
        RxPower = RxOptical.group(1).lstrip(' ')
        print("光衰值："+RxPower+'\n')
    else:
        RxOptical = re.search(r"^epon-onu_\d/\d/\d:\d\s+(.+)$",reply,re.M)
        RxPower = RxOptical.group(1).rstrip('(dbm)')
        print("光衰值：" + RxPower+'\n')


