# -*- coding: utf-8 -*-
import xlrd
import TestExl
import os
from xlutils.filter import process, XLRDReader, XLWTWriter
import xlrd, xlwt
import re
import cchardet
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import time,datetime
from datetime import date,datetime
from xlwt import Workbook,Style

birthdate = "{0}-{1}-{2}"
# /Users/hyshi/Documents/2016-11-18/1曼竜小学/表4参测教师基本信息表.xls
if __name__ == '__main__':
    TestExl.testExl('/Users/hyshi/Documents/2016-11-18/7王家塘小学/表5参测学生基本信息表（五年级甲班）.xls','2016-11-18','newFile')
styleBlueBkg = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: bold on;'); # 80% like
def testExl(readPatch,oldPatch,newPatch):
    print "modif:",readPatch
    rstudent = xlrd.open_workbook(readPatch,formatting_info=True)
    wstudent, style_list = copy2(rstudent)
    try:
        if len(rstudent.sheet_names()) > 0:
            sheet = rstudent.sheet_by_index(0)
            rsheet = wstudent.get_sheet(0)
            if sheet.nrows.numerator > 6:
                info = getIDCardClum(sheet)
                if info == "null":
                    print "info is null"
                    return False
                else:
                    if info['row'] == -1:
                        print "info  row is -1 "
                        return False
                    titleRow = info['row']
                    print titleRow
                    for row in range(titleRow,sheet.nrows.numerator):
                        #获取身份证号
                        if info['idCard'] == -1:
                            print "info  idCard is -1 "
                            return False
                        idcard = str.strip(sheet.cell(row, info['idCard']).value.encode('utf-8'))
                        if len(idcard) < 18:
                            # 如果身份证号没有数据尝试去学籍拿
                            if info['status'] == -1:
                                print "info  status is -1 "
                                return False
                            idcard = str.strip(sheet.cell(row, info['status']).value.encode('utf-8'))
                        if len(idcard) >= 18 and len(idcard) < 40:
                            # 检查学籍长度如果为18 且全是数字则直接写入生日
                            # print idcard
                            group = re.findall('\d{18}|\d{17}\w', idcard)
                            if len(group) > 0:
                                idcard = group[0].upper()
                                writeBirthday(sheet,rsheet,row,info,style_list,rstudent,idcard)
                                # 将学籍列数据写入身份证列
                                writeIdCard(sheet,rsheet,row,info,style_list,rstudent,idcard)
                        if len(idcard)<18 and len(sheet.cell(row,2).value) > 0:

                            styleBlueBkg.font.name =  unicode('宋体', "utf-8");
                            styleBlueBkg.font.bold = False
                            styleBlueBkg.font.charset = styleBlueBkg.font.CHARSET_ANSI_GREEK
                            styleBlueBkg.font.height = 1000
                            rsheet.write(row, info['idCard'], sheet.cell(row, info['idCard']).value, styleBlueBkg)
                            # rsheet.write(row, info['status'], sheet.cell(row, info['status']).value, styleBlueBkg)
                        #如果学籍和身份证都为空 则把学籍和身份证标红

                        #需要获取民族，有的民族只写了前面没有带族针对这个处理
                        checkNation(sheet,rsheet,row,info,style_list)
        else:
            print "readPatch=",readPatch,"sheet size is",len(rstudent.sheet_names())
            # 修改完成后另存
    except NameError:
        print "modify fail path:",readPatch

    newFile = readPatch.replace(oldPatch,newPatch)


    p, f = os.path.split(newFile);
    print  "newFile ===", newFile
    mkdir(p)
    wstudent.save(newFile);
    return True

def checkNation(rsheet,wsheet,row,info,style_list):
    if info['nation'] == -1:
        return
    clum = info['nation']
    xf_index = rsheet.cell_xf_index(row, clum)
    nation = str.strip(rsheet.cell(row, clum).value.encode('utf-8'))
    if '族' not in nation and len(nation) > 0:
        nation += '族'
        print nation,rsheet.cell(row, clum).ctype
        nation = unicode(nation, "utf-8");
        wsheet.write(row, clum, nation, style_list[xf_index])



def writeIdCard(rsheet,wsheet,row,info,style_list,rstudent,idcard):
    if info['idCard'] == -1:
        return
    idcarClum = info['idCard']
    xf_index = rsheet.cell_xf_index(row, idcarClum)
    wsheet.write(row, idcarClum, idcard, style_list[xf_index])

def writeBirthday(rsheet,wsheet,row,info,style_list,rstudent,idcard):
    if info['brithday'] == -1 :
        return
    clum = info['brithday']
    xf_index = rsheet.cell_xf_index(row, clum)
    birthdayfull = str.strip(idcard)[6:-4]
    birthday = birthdate.format(birthdayfull[0:4], birthdayfull[4:6], birthdayfull[-2:])
    #如果类型为日期 则进行日期处理
    # print(rsheet.cell(row, clum).ctype)
    if rsheet.cell(row, clum).ctype == 3:
        date_value = xlrd.xldate_as_tuple(rsheet.cell_value(row, clum), rstudent.datemode)
        date_tmp = date(*date_value[:3]).strftime('%Y-%m-%d')
        # if birthday != date_tmp:
        # sheet.put_cell(row, 4, 3, birthday, 0)
        wsheet.write(row, clum, birthday, style_list[xf_index])
        # print "call styl = 3", row, birthday, date_tmp
    #如果类型为String 则对String 处理
    else:
        # cellbri = str.strip(rsheet.cell(row, clum).value.encode('utf-8'))
        # if birthday != cellbri:
        # sheet.put_cell(row, 4, 1, birthday, 0)
        wsheet.write(row, clum, birthday, style_list[xf_index])
        # print "call styl = 1", row, birthday

##获取单元格中值的类型，类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
def getIDCardClum(sheet):
    info = {'row':-1,'idCard':-1,'status':-1,'brithday':-1,'nation':-1}
    for row in range(sheet.nrows.numerator):
        if row == 5:
            pass
        hasidcard = False
        hasstatus = False
        hasbrith = False
        hasnation = False
        for clum in range(sheet.ncols.numerator):
            if (sheet.cell(row, clum).ctype == 1):
                cell = sheet.cell(row, clum).value.encode('utf-8')
                if '身份证' in str(cell):
                    info['row'] = row + 1
                    info['idCard'] = clum
                    hasidcard = True
                if '学籍' in str(cell):
                    info['row'] = row + 1
                    info['status'] = clum
                    hasstatus = True;
                if '出生' in str(cell) or '日期' in str(cell):
                    info['row'] = row + 1
                    info['brithday'] = clum
                    hasbrith = True;
                if '民族' == str(cell) :
                    info['row'] = row + 1
                    info['nation'] = clum
                    hasnation = True
        if hasidcard and hasstatus and hasbrith and hasnation:
            break
    return info



def modifyExl(readPatch):
    print "patch==",readPatch
    # 打开文件
    rstudent = xlrd.open_workbook(readPatch,formatting_info=True)

    wstudent, style_list = copy2(rstudent)

    if len(rstudent.sheet_names()) > 0:
        sheet = rstudent.sheet_by_index(0)
        wtsheet = wstudent.get_sheet(0)
        cell_info = sheet.cell(6, 4).value
        birthday = str.strip(sheet.cell(6, 3).value.encode('utf-8'))[6:-4]
        birthdate = "{0}{1}-{2}"
        # print  birthdate.format(birthday[0:4],birthday[4:6],birthday[-2:])
        # print  cell_info
        # print  sheet.name,sheet.nrows,sheet.ncols
        # print  sheet.nrows.numerator
        for row in range(6, sheet.nrows):
            # newBirt = birthdate.format(birthday[0:4], birthday[4:6], birthday[-2:])
            # oldBirt = str.strip(sheet.cell(row, 4).value.encode('utf-8'))

            xf_index = sheet.cell_xf_index(row, 4)
            wtsheet.write(row, 4, "ddd", style_list[xf_index])
            # print "newBirt{},olde:{}",newBirt,oldBirt
            # if newBirt != oldBirt:
            #     sheet.put_cell(row, 4, 1, newBirt, 0)
            # newWs = wstudent.get_sheet(0);
            # newWs.write(row, 4, newBirt);
            # print sheet.cell(row,4)
            # print  "{}{}",str.strip(sheet.cell(row,3).value.encode('utf-8')),str.strip(sheet.cell(row,4).value.encode('utf-8'))

    #修改完成后另存

    wstudent.save('/Users/hyshi/Documents/2016-11-18/text.xls');

def mkdir(path):
    # 引入模块
    import os

    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # print path + ' 创建成功'
        # 创建目录操作函数
        os.makedirs(path)
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        # print path + ' 目录已存在'
        return False
def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb, 'unknown.xls'),
        w
    )
    return w.output[0][1], w.style_list
