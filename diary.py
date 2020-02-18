#!/usr/bin/python3
#!/usr/bin/env python3

import os
import re
import sys
import time
import getopt
import socket
from ftplib import FTP
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
import configparser
from datetime import date,datetime,time

############################
# @berif 打印 help 信息
def usage( ):
    print(' Usage: diary <cmd>')
    print(' cmd:')
    print('     commit  ：提交一条记录，Enter 结束输入')
    print('     push    : 向 FTP 服务器推送本地记录文件')
    print(' Diary v1.0.0  2020/2/3 ( leoyang20102013@163.com )')

############
## ini 文件操作
def construct_default_config( config_file_name):
    server_ip_str = input(" 服务器IP : ")

    if not re.match(  r'^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$' ,server_ip_str):
        print ( "IP address invaild ")
        exit()

    server_port_str = input(" 服务器Port : ")
    if not re.match(  r'^(?:[0-9]{1,5})$' ,server_port_str):
        print ( "Port invaild ")
        exit()

    server_online_str = input(" 网络支持 (y/n): ")
    if server_online_str != "y" :
        server_online_str = "n"

    server_user_str = input(" 用户名: ")
    server_pass_str = input(" 用户密码: ")
    local_file_name_str = input(" 日志文件名: ")

    config = configparser.ConfigParser()
    # 默认，每周第一天第一次记录时，上传日志
    config["DEFAULT"] = { "auto_upload":"1" , "local_file_name":local_file_name_str  }
    auto_upload_int = 1
    local_file_name = local_file_name_str

    config["FTP Server"] = {}
    config["FTP Server"]["online"] = server_online_str
    config["FTP Server"]["serverIP"] = server_ip_str
    config["FTP Server"]["serverPort"] = server_port_str
    config["FTP Server"]["userName"] = server_user_str
    config["FTP Server"]["password"] = server_pass_str
    with open(  sys.path[0] + '/'+config_file_name ,"w") as fd :
        config.write( fd )  
    return auto_upload_int,  local_file_name , server_ip_str,server_port_str,server_online_str,server_user_str,server_pass_str

## 测试：test_parser_config()
def parser_config( config_file_name="config.ini" ):
    print( " *****  config_file_name : " ,config_file_name )

    if config_file_name == "":
        return None
    else :
        if not config_file_name.endswith(".ini") :
            return None

    config_file_path  = sys.path[0] +'/'+config_file_name
    if not os.access( config_file_path , os.F_OK ):
        print (" 首次使用,配置服务器参数：")
        return construct_default_config( config_file_name )
    else :
        cfg_file = configparser.ConfigParser()
        cfg_file.read( config_file_path)

        if "DEFAULT" in cfg_file :
            auto_upload_int = cfg_file["DEFAULT"].getint("auto_upload")
            local_file_name  = cfg_file["DEFAULT"]["local_file_name"]
        else :
            auto_upload_int = 1
            local_file_name="local_diary.xlsx"

        if "FTP Server" in cfg_file :
            server_ip_str = cfg_file["FTP Server"]["serverIP"]
            server_port_str = cfg_file["FTP Server"]["serverPort"]
            server_online_str = cfg_file["FTP Server"]["online"]
            server_user_str = cfg_file["FTP Server"]["userName"]
            server_pass_str = cfg_file["FTP Server"]["password"]
        else :
            server_ip_str = "192.168.1.105"
            server_port_str = "21"
            server_online_str= "n"
            server_user_str = "user"
            server_pass_str = "123"
    return auto_upload_int , local_file_name , server_ip_str,server_port_str,server_online_str,server_user_str,server_pass_str

def get_week_date():
    datetime_obj = datetime.now()
    cur_date_str = datetime_obj.date().strftime("%Y-%m-%d")
    cur_time_str = datetime_obj.time().strftime("%H:%M:%S")
    cur_week_in_year_tuple = datetime_obj.isocalendar()
    return cur_date_str,cur_time_str,cur_week_in_year_tuple[1],cur_week_in_year_tuple[2]

## 测试: test_diary.test_get_workbook()
def get_workbook( fileName = "yangj_log.xlsx" ):
    if fileName == "":
        return None
    else :
        if not fileName.endswith(".xlsx")  :
            return None

    if not os.access( fileName , os.F_OK ):
        wb = openpyxl.Workbook()
        sheet = wb .create_sheet("temp")
        wb.save(fileName)
        print (" File not found ,then create new workbook")
        return wb
    print ( " Loading file...%s" % fileName )
    return openpyxl.load_workbook( fileName )

def get_sheet( file,wb_obj , title_str ):
    try:
        sheet_list = wb_obj.get_sheet_names()
        sheet_index = sheet_list.index( title_str )
    except ValueError :
        temp_sheet_index = sheet_list.index( "temp" )
        sheet_obj = wb_obj.copy_worksheet ( wb_obj.worksheets[ temp_sheet_index ]  )
        sheet_obj.title = title_str
        wb_obj.save( file )
    else :
        sheet_obj = wb_obj.worksheets[ sheet_index ]
    finally:
        return sheet_obj

def ftp_upload( remote_file , local_file , ip,  userName, password ):
    try:
        ftp = FTP(host=ip , user=userName , passwd=password)
    except (socket.error, socket.gaierror):
        print ("异常：服务器连接失败")
        return None
    
    ftp.login( user=userName , passwd=password)
    print ( "登陆成功:  %s" % ftp.getwelcome())
    # 匿名登陆，文件写入 ftp-anonymous/ 
    # ftp = FTP(host="192.168.1.105")
    # ftp.login()
    with open( local_file,"rb") as fd :
        ftp.storbinary('STOR ' + remote_file , fd ,1024 )
    ftp.quit()

#################
def main():

    if sys.argv.__len__() <= 1:
        usage()
        exit()

    date_str,time_str,week_int,weekDay_int = get_week_date()
    print ("%s      第%d周 第%d天" % ( date_str, week_int , weekDay_int ) )

    auto_upload_int, local_file_name, server_ip_str,server_port_str,server_online_str,server_user_str,server_pass_str = parser_config("config.ini") 
    local_file_path  = sys.path[0] +'/' + local_file_name

    if sys.argv[1] == "commit" :
        content_str = input("[ 随手记 ]")
        if content_str != "":
            wb = get_workbook(  local_file_path )
            sheet = get_sheet( local_file_path , wb , str( week_int ) )
            privous_str = sheet.cell( weekDay_int+1 ,2).value
            sheet.cell( weekDay_int+1 ,2).alignment = Alignment( horizontal="left",vertical="top")
            
            if privous_str == None :
                # 空白cell
                sheet.cell( weekDay_int+1 ,2).value = "[ "+ time_str+"] "+ content_str
            else :
                sheet.cell( weekDay_int+1 ,2).value = privous_str +"\n[ "+ time_str+"]  "+ content_str
            wb.save( local_file_path )
            
            if  (auto_upload_int == weekDay_int) and (server_online_str == "y")  :
                ret = ftp_upload( local_file_name , local_file_path ,server_ip_str, server_user_str, server_pass_str )
                if ret == None :
                    print ("确认服务器IP配置正确或服务器已经启动")
        else :
            print (" 输入内容为空，不写入任何内容")
            exit()
    elif (sys.argv[1] == "push") and (server_online_str == "y") :
        ret = ftp_upload( local_file_name , local_file_path ,server_ip_str, server_user_str, server_pass_str )
        if ret == None :
            print ("确认服务器IP配置正确或服务器已经启动")
        exit()
    elif sys.argv[1] == "help" :
        usage( )
        exit()
    else:
        exit()
    return 

#########
if __name__ == "__main__":
    main()