# -*- coding: utf-8 -*- 
from xlrd import open_workbook
from xlutils.copy import copy
import os
import xlrd
import xlwt

filePath = "F:"
fileName = "connect.xls"
fileCatalog = filePath + "\\" + fileName
curr_row = 1

def makeDict(curr_name, curr_telephone):
    curr_dict = {}
    curr_dict['name'] = curr_name
    curr_dict['telephone'] = curr_telephone 
    return curr_dict

def operatorExcel(name, telephone):

    #如果没有该excel，则创建excel
    if os.path.isfile(fileCatalog) is False:
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet('personinfo')
        sheet.write(0, 0, 'Name')
        sheet.write(0, 2, 'Telephone')
        wbk.save(fileCatalog)
    
    rb = open_workbook(fileCatalog)
    rs = rb.sheet_by_name(u'personinfo')#通过名称获取sheet

    num_rows = rs.nrows#获取row行数
    #print 'num_rows:' + str(num_rows)
    rscurr_row = 1 #0为说明行
    
    wb = copy(rb)
    ws = wb.get_sheet(0)
    wscurr_row = 0
    

    ws.write(wscurr_row, 0, 'Name')
    ws.write(wscurr_row, 2, 'Telephone')
    wscurr_row += 1
    
    if rscurr_row == num_rows:
        ws.write(wscurr_row, 0, name)
        ws.write(wscurr_row, 2, telephone)
        wb.save(fileCatalog)
        return 3  #表示没有相同数据
    
    flag_same = 0 #0: 不同名，1：数据相同， 2： 名字存在，号码不确定 3: 已处理
    while rscurr_row < num_rows:
        arr_name = rs.cell_value(rscurr_row, 0)#代表当前数组的name，因为其余为空的
        curr_telephone = rs.cell_value(rscurr_row, 2)
        
        temp_arr = []
        temp_dict = makeDict(arr_name, curr_telephone)
        temp_arr.append(temp_dict)

        if arr_name == name:
            if curr_telephone == telephone:
                flag_same = 1
            else:
                flag_same = 2
            
        rscurr_row += 1
        while True:
            if rscurr_row >= num_rows:
                break
            
            curr_name = rs.cell_value(rscurr_row, 0)
            curr_telephone = rs.cell_value(rscurr_row, 2)
            
            if flag_same == 2:
                if curr_telephone == telephone:
                    flag_same = 1
                    
            if curr_name.strip() == '' or curr_name == arr_name:
                temp_dict = makeDict(arr_name, curr_telephone)
                temp_arr.append(temp_dict)
                rscurr_row += 1
            else:
                break
            
        if flag_same == 2:
            temp_dict = makeDict(arr_name, telephone)
            temp_arr.append(temp_dict)
            flag_same = 3
            
        flag_first = True
        #print temp_arr
        for arr_iter in range(0, len(temp_arr)):
            ws.write(wscurr_row, 0, '')
            temp_dict = temp_arr[arr_iter]
            #print temp_dict
            if flag_first is True:
                ws.write(wscurr_row, 0, temp_dict['name'])
                flag_first = False
            ws.write(wscurr_row, 2, temp_dict['telephone'])
            wscurr_row += 1

    if flag_same == 0:
        ws.write(wscurr_row, 0, name)
        ws.write(wscurr_row, 2, telephone)
        flag_same = 3
        
    wb.save(fileCatalog)
    return flag_same

if __name__ == '__main__':
    cnt = 1
    #设置路径
    filePath = raw_input('Enter file path:(like: C:\\administor)\n')
    #filePath 
    #设置文件名
    while True:
        group = raw_input('Enter gz: 高中 dx:大学 gs:公司\n')
        if group == 'gz' or group == 'dz' or group == 'gs':
            fileCatalog = filePath + "\\" + group + "_" + fileName
            print fileCatalog
            break
        else:
            print 'Please Enter right name!!!'
    while True:
        print '第' , cnt , '组:'
        cnt += 1
        name = raw_input('Enter Name:')
        name = name.decode('gbk').encode('utf-8')
        if name.strip() == '':
            break
        telephone = raw_input('Telephone:')
        if operatorExcel(name, telephone) == 3:
            print '添加成功!'
        else:
            print '信息已存在!!!'
        print ''
