# -*- coding: utf-8 -*-

from xlwt import *
import sys
import re

def transformTxt2Xls(para):
    totalString = ''
    tempString = ''
    tempStringList = []
    glassModel = ''
    detectTime = ''
    shapeRangeAB = 0.0
    shapeRangeCD = 0.0
    pointCount = []     #pointCount对应AB/BC/CD/DA边的点数量
    
#解析原始文件目录和输出文件目录
    srcTxtFileName = ''
    dstXlsFlieName = ''
    tempStringList = para.split('*#*')
    if len(tempStringList) != 2:
        srcTxtFileName = 'tempExportFile.txt'
        dstXlsFlieName = 'defaultXLS.xls'
    else:
        srcTxtFileName = tempStringList[0]
        dstXlsFlieName = tempStringList[1]
        if srcTxtFileName == '' or dstXlsFlieName == '':
            srcTxtFileName = 'tempExportFile.txt'
            dstXlsFlieName = 'defaultXLS.xls'			
     
    #读取和解析原始数据 
    with open(srcTxtFileName, 'r') as f:
        totalString = f.read()
        #print(totalString)
    print(f.closed)
    
    #得到玻璃型号和导出日期
    matchResult = re.search('glassModel:.+\n',  totalString)
    if matchResult == None:
        print 'shapeRange Format Error'
        return -1
    glassModel = matchResult.group().split(':')[1]
    matchResult = re.search('detectTime:.+\n',  totalString)
    if matchResult == None:
        print 'detectTime Format Error'
        return -1
    detectTime = matchResult.group().split(':')[1]
    
    #得到数据的合格尺寸范围
    matchResult = re.search('shapeRange:.+\n',  totalString)
    if matchResult == None:
        print 'shapeRange Format Error'
        return -1
    tempString = matchResult.group().split(':')[1]
    tempStringList = tempString.split(',')
    if len(tempStringList) != 2:
        print 'shapeRange Format Error'
        return -1
    shapeRangeAB = float(tempStringList[0])
    shapeRangeCD = float(tempStringList[1])
    print shapeRangeAB,  shapeRangeCD
    
    #得到4个边的点数量
    matchResult = re.search('pointCount:.+\n',  totalString)
    if matchResult == None:
        print 'pointCount Format Error'
        return -1
    tempString = matchResult.group().split(':')[1]
    tempStringList = tempString.split(',')
    if len(tempStringList) != 4:
        print 'pointCount Format Error'
        return -1
    my_list = [0, 1, 2, 3]
    for i in my_list:
        pointCount.append((int)(tempStringList[i]))
    print pointCount[0],  pointCount[1], pointCount[2], pointCount[3]
    
    #解析尺寸偏差数据
    rowStringList = totalString.split('\n')
    indexList = []      #记录所有有效数据的行数索引
    sideName = ['AB',  'BC',  'CD',  'DA']
    for i in range(0,  len(rowStringList)):
        tempString = rowStringList[i]
        if tempString.find('*#*') == -1:   
            continue    #跳过表头或注释文字
        sideList = tempString.split('*#*')
        if len(sideList) != 5:
            continue    #跳过测量数据缺失的记录
        #依次判断每条边的点数是否与文件中的配置一致
        error = ''
        for j in range(1,  5):
            if len(sideList[j].split(':')) != pointCount[j - 1]:
                error = str(i + 1) + ',' + sideName[j - 1]
                print 'pointCount Error: ' + error
                break;
        if error == '':
            indexList.append(i)

    #开始将数据导入xls文件   
    book = Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet1') #创建一个sheet

    #创建一个样式----------------------------
    styleNG = XFStyle()
#    pattern = Pattern()
#    pattern.pattern = Pattern.SOLID_PATTERN
#    pattern.pattern_fore_colour = Style.colour_map['yellow'] #设置单元格背景色为黄色
#    styleNG.pattern = pattern  
    fontNG = Font()         # 为样式创建字体
    fontNG.colour_index = 2     #设置字体颜色为红色
    styleNG.font = fontNG
    
    
    #插入表头
    rowIndex = 0
    colIndex = 0
    sheet.write(rowIndex, 0, label = '检测设备:')
    sheet.write(rowIndex, 1, label = 'UCL100')
    sheet.write(rowIndex, 3, label = '导出日期:')
    sheet.write(rowIndex, 4, label = detectTime)
    sheet.write(rowIndex, 6, label = '玻璃型号:')
    sheet.write(rowIndex, 7, label = glassModel)
    
    rowIndex += 1
    sheet.write(rowIndex, 0, label = '产线:')
    sheet.write(rowIndex, 1, label = '1号线')
    sheet.write(rowIndex, 3, label = '地点:')
    sheet.write(rowIndex, 4, label = '上海')
    sheet.write(rowIndex, 6, label = '取点间隔:')
    sheet.write(rowIndex, 7, label = '100mm')
    
    rowIndex += 1
    sheet.write(rowIndex, 0, label = '基准边:')
    sheet.write(rowIndex, 1, label = 'CD')
    sheet.write(rowIndex, 2, label = 'DA')
    sheet.write(rowIndex, 3, label = '非基准边:')
    sheet.write(rowIndex, 4, label = 'AB')
    sheet.write(rowIndex, 5, label = 'BC')
    
    rowIndex += 1
    sheet.write(rowIndex, 0, label = '要求数值:')
    sheet.write(rowIndex, 1, label = str(shapeRangeCD) + 'mm')
    sheet.write(rowIndex, 2, label = str(shapeRangeCD) + 'mm')
    sheet.write(rowIndex, 4, label = str(shapeRangeAB) + 'mm')
    sheet.write(rowIndex, 5, label = str(shapeRangeAB) + 'mm')
   
    rowIndex += 2 
    sheet.write(rowIndex, 0, label = '检测时间')
    colIndex = 1
    for col in range(1,  pointCount[0]):
        sheet.write(rowIndex, colIndex, label = 'AB边尺寸点' + str(col))
        colIndex += 1
    for col in range(1,  pointCount[1]):
        sheet.write(rowIndex, colIndex, label = 'BC边尺寸点' + str(col))
        colIndex += 1
    for col in range(1,  pointCount[2]):
        sheet.write(rowIndex, colIndex, label = 'CD边尺寸点' + str(col))
        colIndex += 1
    for col in range(1,  pointCount[3]):
        sheet.write(rowIndex, colIndex, label = 'DA边尺寸点' + str(col))
        colIndex += 1
    
    
    #插入测量数据
    rowIndex += 1
    for row in indexList:
        tempString = rowStringList[row]
        strList = tempString.split('*#*')
        colIndex = 0
        #插入每条记录的检测时间----改格式 不能有中文
        sheet.write(rowIndex, colIndex, label = strList[0])
        colIndex += 1
        #先计算AB边
        sideStrList = strList[1].split(':')
        tempFloatValue = 0.0
        for col in range(1,  pointCount[0]):
            tempFloatValue = float('%0.4f'%float(sideStrList[col])) #设置导出数值精度为4位小数
            if abs(tempFloatValue) > shapeRangeAB:
                sheet.write(rowIndex, colIndex, tempFloatValue, style=styleNG)
            else:
                sheet.write(rowIndex, colIndex, tempFloatValue)
            colIndex += 1
         
        #计算BC边
        sideStrList = strList[2].split(':')
        for col in range(1,  pointCount[1]):    #从每条边的第二个数据开始输出
            tempFloatValue = float('%0.4f'%float(sideStrList[col])) #设置导出数值精度为4位小数
            if abs(tempFloatValue) > shapeRangeAB:
                sheet.write(rowIndex, colIndex, tempFloatValue, style=styleNG)
            else:
                sheet.write(rowIndex, colIndex, tempFloatValue)
            colIndex += 1  
        
        #计算CD边
        sideStrList = strList[3].split(':')
        for col in range(1,  pointCount[2]):
            tempFloatValue = float('%0.4f'%float(sideStrList[col])) #设置导出数值精度为4位小数
            if abs(tempFloatValue) > shapeRangeCD:
                sheet.write(rowIndex, colIndex, tempFloatValue, style=styleNG)
            else:
                sheet.write(rowIndex, colIndex, tempFloatValue)
            colIndex += 1   
        
        #计算DA边
        sideStrList = strList[4].split(':')
        for col in range(1,  pointCount[3]):
            tempFloatValue = float('%0.4f'%float(sideStrList[col])) #设置导出数值精度为4位小数
            if abs(tempFloatValue) > shapeRangeCD:
                sheet.write(rowIndex, colIndex, tempFloatValue, style=styleNG)
            else:
                sheet.write(rowIndex, colIndex, tempFloatValue)
            colIndex += 1   
        rowIndex += 1
        
    book.save(dstXlsFlieName)
        
def excel_write(data_list):
    book = Workbook(encoding='utf-8')
    sheet = book.add_sheet('Sheet1') #创建一个sheet

    #创建一个样式----------------------------
    style = XFStyle()
    pattern = Pattern()
    pattern.pattern = Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = Style.colour_map['yellow'] #设置单元格背景色为黄色
    style.pattern = pattern
    #-----------------------------------------

    sheet.write(0, 0, label = 'ICAO') #给第0行的第1列插入值
    sheet.write(0, 1, label = 'Location') #给第0行的第2列插入值
    sheet.write(0, 2, label='Airport_Name')
    sheet.write(0, 3, label='Country')

    #循环插入值
    for num,x in enumerate(data_list):
        index=num+1
        print(index)
        # if index!=0:
        sheet.write(index, 0, label = x["ICAO"])
        sheet.write(index, 1, label = x["Location"])
        sheet.write(index, 2, label=x["Airport_Name"], style=style) #将样式添加到此单元格
        sheet.write(index, 3, label=x["Country"])
    book.save('air.xls')
#测试数据
#data=[{"ICAO": "DSG", "Location": "SDGSDG", "Airport_Name": "sdgsdg??sdg",
#     "Country": "dfsdg"},{"ICAO": "DSG", "Location": "SDGS23G", "Airport_Name": "sdgsdg23??sdg",
#     "Country": "354746"}]
#excel_write(data)

#测试传参
if len(sys.argv) < 2:
    print 'argv Error'
else:
    para = sys.argv[1]
    transformTxt2Xls(para)


