import xlrd
import pandas as pd
import re
import telnetlib
import time
import os

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




#从光衰表中获取OLT的IP，设备型号Device_ID，PON口以及ONU_ID
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
    user = 'lyread'
    password = 'read1234'
    commands = []
    if Device == 'MA5680T':
        commands = [user,password,'enable','config','interface epon ' + SLOT + '/' + BOARD,'display ont optical-info ' + PORT + ' ' + ONU_ID]
    elif Device == 'MA5683T':
        commands = [user, password,'enable','config','interface epon ' +SLOT + '/' + BOARD,'display ont optical-info ' + PORT + ' ' + ONU_ID]
    elif Device == 'C300':
        commands = [user,password,'show pon power onu-rx epon-onu_'+PON+':'+ONU_ID]
    elif Device == 'C220':#命令不对
        commands = [user, password, 'show pon power onu-rx epon-onu_' + PON + ':' + ONU_ID]
    else:
        print("错误！！！没有这种设备，联系57591添加命令")
    return commands


def telnet():
    try:
        tn = telnetlib.Telnet(host=IP, port=23,timeout=10)   # 尝试开启Telnet Socket
    except ConnectionRefusedError:
        print("错误！！！目标计算机连接超时，请检查链路是否正常！")
    else:
        commands = command_make(Device=DEVICE_ID)            # 根据设备型号制作命令
        print('努力查询中，请稍后......\n')
        # 输入命令并执行
        for command in commands:
            time.sleep(0.5)
            write_command(tn,command)
        time.sleep(2)
        reply = (str(tn.read_very_eager(),encoding='utf-8'))
        state = re.search(r"^\s+Failure:(.+)$",reply,re.M)
        error_code = state.group(1).lstrip(' ')
        print(error_code)
        if error_code == "The ONT is not online":
            print("光猫未上线\n")
        else:
            sample = reply.replace('(', '').replace(')', '')
            if DEVICE_ID == 'MA5680T':
                RxOptical = re.search(r"^\s+Rx optical powerdBm\s+:(.+)$", sample, re.M)
                RxPower = RxOptical.group(1).lstrip(' ')
                print("光衰值："+RxPower+'\n')
            elif DEVICE_ID == 'MA5683T':
                RxOptical = re.search(r"^\s+Rx optical powerdBm\s+:(.+)$", sample, re.M)
                RxPower = RxOptical.group(1).lstrip(' ')
                print("光衰值：" + RxPower + '\n')
            else:
                RxOptical = re.search(r"^epon-onu_\d/\d/\d:\d{1,3}\s+(.+)$",reply,re.M)
                RxPower = RxOptical.group(1).replace('(dbm)','')
                print("光衰值：" + RxPower+'\n')
        tn.close()


def scan_files(directory, prefix=None, postfix=None):
    files_list = []
    for root, sub_dirs, files in os.walk(directory):
        for special_file in files:
            if postfix:
                if special_file.endswith(postfix):
                    files_list.append(os.path.join(root, special_file))
            elif prefix:
                if special_file.startswith(prefix):
                    files_list.append(os.path.join(root, special_file))
            else:
                files_list.append(os.path.join(root, special_file))
    return files_list