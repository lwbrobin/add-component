# -*- coding: utf-8 -*-
import xlrd
import xlwt

class componet:
    def __init__(self, name, type, len):
        self.name = name    #构件名
        self.type = type    #截面类型
        self.len = len      #构件长度

    # 打印构件名和长度
    def disp(self):
         print("name:",self.name,"type:",self.type, "length:", self.len)

    # 得到构件的厚度
    def get_thick(self):
        if(self.type.startswith('PL')):
            tmp = int(self.type.find('*'))
            tmpstr = self.type[2:tmp]
            return int(tmpstr)
            #return int(self.type[2,self.type.find('*')])
        elif(self.type.startswith('L')):
            tmp = int(self.type.find('*'))
            tmpstr = self.type[tmp+1:]
            return int(tmpstr)
        else:
            return 0

    # 得到构件的宽度
    def get_width(self):
        if(self.type.startswith('PL')):
            tmp = int(self.type.find('*'))
            tmpstr = self.type[tmp+1:]
            return int(tmpstr)
            #return int(self.type[2,self.type.find('*')])
        elif(self.type.startswith('L')):
            tmp = int(self.type.find('*'))
            tmpstr = self.type[1:tmp]
            return int(tmpstr)
        else:
            return 0


def write_excel_xls(value, position):
    comnum = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet('sheet1')  # 在工作簿中新建一个表格
    sheet.write(0, 0, '构件名')    # 向表格中写入数据（对应的行和列）
    sheet.write(0, 1, '宽度')
    sheet.write(0, 2, '截面类型')
    sheet.write(0, 3, '厚度')
    sheet.write(0, 4, '长度')
    sheet.write(0, 5, '总长')
    rownum = 1
    for i in range(0, comnum):
        total_len = 0.0
        for j in range(0, len(value[i])):
            sheet.write(rownum, 0, value[i][j].name)
            sheet.write(rownum, 1, value[i][j].get_width())
            sheet.write(rownum, 2, value[i][j].type)
            sheet.write(rownum, 3, value[i][j].get_thick())
            sheet.write(rownum, 4, value[i][j].len)
            total_len += value[i][j].len
            rownum +=1
        sheet.write(rownum, 5, total_len/1000)
        rownum += 1
    workbook.save(position)  # 保存工作簿路径
    print("xls格式表格写入数据成功！")

def find_vec(com, tar, result, err_control): #传入一组componet，目标长度，结果，控制余量
    if(not com or tar < com[-1].len):
        return False

    for ind in range(len(com)):
        remain = tar - com[ind].len
        if(remain <= 0):
            continue
        if(remain < err_control):
            result.append(com[ind])
            return True

        tmp = []
        if(find_vec(com[ind+1:], remain, tmp, err_control)):
            result.append(com[ind])
            result.extend(tmp)
            return True

    return False

#文件位置
ExcelFile = xlrd.open_workbook(r'D:\test.xls')
#获取目标EXCEL文件sheet名
#print(ExcelFile.sheet_names())
sheet1=ExcelFile.sheet_by_name('Sheet1')    #读取表格，默认表格名为Sheet1

#打印sheet的名称，行数，列数
print(sheet1.name,sheet1.nrows,sheet1.ncols)

#获取整行或者整列的值
components = []

#读取所有构件
for ind in range(1, sheet1.nrows):
    a = componet(sheet1.cell_value(ind, 0), sheet1.cell_value(ind, 1),sheet1.cell_value(ind, 2))
    if(a.get_thick() == 16):  #板厚控制
        components.append(a)

#把构件按长度排序
components.sort(key=lambda x:x.len)
components.reverse()

result = []
control_up = 10000      #长度上限
err_control = 100
while(len(components) > 1):
    tmp_res = []
    total_len = 0.0
    for eve in range(len(components)):
        total_len += components[eve].len
    if(total_len < control_up):   #如果总长都小于给定值，则没有必要继续循环找，直接返回
        tmp_res.extend(components)
        result.append(tmp_res)
        components.clear()
        continue
    while(not find_vec(components, control_up, tmp_res, err_control)):
        err_control += 100

    result.append(tmp_res)
    for val in tmp_res:
        components.remove(val)

write_excel_xls(result,"D:\\result.xls")  #输出结果到H:\result.xls

