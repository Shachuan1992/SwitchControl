import xlwt
import xlrd


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

write_excel('test.xls',0,0,'Hello World')
