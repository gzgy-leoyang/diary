import openpyxl
from openpyxl import Workbook

import os
import sys
import getopt
import time
from datetime import date,datetime,time

############################
# @berif 打印 help 信息
def usage( ):
    print(' Usage: python3 diary.py  <-o context>')
    print(' options:')    
    print(' -o')
    print(' Diary v1.0.0  2020/2/3 ( leoyang20102013@163.com )')

###################
# @ berif
#
def parser_content( argv ):
    if argv[1] == "commit" :
        return input(" 记录工作内容：")
    elif (argv[1] == "--help") or argv[1] == "-h" :
        usage( )
        exit()
    else:
        exit()
    
def get_week_date():
    datetime_obj = datetime.now()
    cur_date_str = datetime_obj.date().strftime("%Y-%m-%d")
    cur_time_str = datetime_obj.time().strftime("%H:%M:%S")
    cur_week_in_year_tuple = datetime_obj.isocalendar()
    # print (" %s %s %d %d" % ( cur_date_str, cur_time_str,cur_week_in_year_tuple[1],cur_week_in_year_tuple[2]) )

    cur_week_int = cur_week_in_year_tuple[1]
    cur_weekDay_int = cur_week_in_year_tuple[2]
    return cur_date_str,cur_time_str,cur_week_int,cur_weekDay_int

def get_workbook( fileName ):
    # ## 查询是否有文件，如果有该文件，执行删除，再重新建同名文件  
    global file_name

    cur_path  = sys.path[0]
    file_name  = cur_path+'/'+fileName
    try:
        if os.access( file_name , os.F_OK ):
            print ( 'diary file .... OK' )
    except FileNotFoundError as reason:
        print (" FileNotFoundError " , reason )
        exit()
    else :
        print (" File is existed " )
    finally:
        wb = openpyxl.load_workbook( file_name )
        print (" Return workbook object ,number of sheet :" , len ( wb.get_sheet_names()))
        return wb

def get_sheet( wb_obj , title_str ):
    global file_name
    try:
        sheet_list = wb_obj.get_sheet_names()
        sheet_index = sheet_list.index( title_str )
    except ValueError :
        temp_sheet_index = sheet_list.index( "temp" )
        print ( "temp_sheet_index : ",temp_sheet_index  )
        sheet_obj = wb_obj.copy_worksheet ( wb_obj.worksheets[ temp_sheet_index ]  )
        print ( "temp_sheet_index ..2")
        
        sheet_obj.title = title_str
        print ( "temp_sheet_index ..3")

        wb_obj.save( file_name )
        print ( " Sheet has NOT been found , create new sheet ")
    else :
        print ( " Sheet has been found ")
        sheet_obj = wb_obj.worksheets[ sheet_index ]
    finally:
        print ( " Sheet ready")
        return sheet_obj
# append context in the sheet

# 日记文件名
diary_file = "log.xlsx"

# 包含路径的日记文件名
file_name = None

def main():
    global file_name
    # 获取当前时间，包括年周数（一年内的第几周）和周日数（一周内的第几天）
    date_str,time_str,week_int,weekDay_int = get_week_date()
    print (" %s  %s  %dth WEEK %dth Day " % ( date_str,time_str,week_int,weekDay_int ))
    
    # 根据年周数，确定表格名称
    sheet_title_str = str( week_int )

    # 提取命令行参数中的日志写入内容
    content_str = parser_content( sys.argv )
    print ("content_str: ",content_str )

    # 获取工作簿及其中的对应表
    wb = get_workbook(  diary_file )
    sheet = get_sheet( wb , sheet_title_str )
    print ( "c[2,1]:" ,sheet.cell(2,1).value )
    
    # 添加内容 
    new_str = sheet.cell( weekDay_int+1 ,2).value +"\n[ "+ time_str+"]"+ content_str
    sheet.cell( weekDay_int+1 ,2).value = new_str
    wb.save(file_name)
    return 

#################
## <<  程序入口 >> ##
if __name__ == "__main__":
    main()