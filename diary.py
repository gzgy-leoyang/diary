import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment

from ftplib import FTP

import os
import sys
import getopt
import time
from datetime import date,datetime,time

import configparser

#################################################33
# 日记文件名
diary_file = "log.xlsx"

# 包含路径的日记文件名
file_name = None

# 全局，当前路径
cur_path = None

auto_upload_int = None
server_ip_str = None
server_port_str = None
server_user_str = None
server_pass_str = None

############################
# @berif 打印 help 信息
def usage( ):
    print(' Usage: python3 diary.py commit ')
    print(' Diary v1.0.0  2020/2/3 ( leoyang20102013@163.com )')

def construct_default_config():

    config = configparser.ConfigParser()
    # 默认，每周第一天第一次记录时，上传日志
    config["DEFAULT"] = { "auto_upload":"1"}
    auto_upload_int = 1
    server_ip_str = input(" 服务器IP (xxx.xxx.xxx.xxx): ")
    server_port_str = input(" 服务器Port (0~65535): ")
    server_user_str = input(" 用户名: ")
    server_pass_str = input(" 用户密码: ")

    config["FTP Server"] = {}
    config["FTP Server"]["serverIP"] = server_ip_str
    config["FTP Server"]["serverPort"] = server_port_str
    config["FTP Server"]["userName"] = server_user_str
    config["FTP Server"]["password"] = server_pass_str

    with open("config.ini","w") as config_file :
        config.write(config_file)
    
    return auto_upload_int,server_ip_str,server_port_str,server_user_str,server_pass_str

def parser_config():
    cfg_name  = cur_path+'/config.ini'
    if not os.access( cfg_name , os.F_OK ):
        print (" 首次使用需要配置服务器参数：")
        return construct_default_config()
    else :
        cfg_file = configparser.ConfigParser()
        cfg_file.read( cfg_name)
        if "DEFAULT" in cfg_file :
            auto_upload_int = cfg_file["DEFAULT"].getint("auto_upload")
        else :
            auto_upload_int = 1

        if "FTP Server" in cfg_file :
            server_ip_str = cfg_file["FTP Server"]["serverIP"]
            server_port_str = cfg_file["FTP Server"]["serverPort"]
            server_user_str = cfg_file["FTP Server"]["userName"]
            server_pass_str = cfg_file["FTP Server"]["password"]
        else :
            server_ip_str = "192.168.1.105"
            server_port_str = "21"
            server_user_str = "user"
            server_pass_str = "111"
    return auto_upload_int,server_ip_str,server_port_str,server_user_str,server_pass_str


###################
# @ berif
#
def parser_content( argv ):
    if argv[1] == "commit" :
        # 返回记录的内容
        return input(" [随手记 ]")  
    elif argv[1] == "push":
        # ftp_service()
        exit()
    elif (argv[1] == "help") :
        usage( )
        exit()
    else:
        exit()
    
def get_week_date():
    datetime_obj = datetime.now()
    cur_date_str = datetime_obj.date().strftime("%Y-%m-%d")
    cur_time_str = datetime_obj.time().strftime("%H:%M:%S")
    cur_week_in_year_tuple = datetime_obj.isocalendar()
    print (" %s      全年第%d周 本周第%d天" % ( cur_date_str,cur_week_in_year_tuple[1],cur_week_in_year_tuple[2]) )

    cur_week_int = cur_week_in_year_tuple[1]
    cur_weekDay_int = cur_week_in_year_tuple[2]
    return cur_date_str,cur_time_str,cur_week_int,cur_weekDay_int

def get_workbook( fileName ):
    # ## 查询是否有文件，如果有该文件，执行删除，再重新建同名文件  
    global file_name
    # cur_path  = sys.path[0]
    # file_name  = cur_path+'/'+fileName
    if not os.access( file_name , os.F_OK ):
        print (" FileNotFoundError " , reason )
        exit()
    return openpyxl.load_workbook( file_name )

def get_sheet( wb_obj , title_str ):
    global file_name
    try:
        sheet_list = wb_obj.get_sheet_names()
        sheet_index = sheet_list.index( title_str )
    except ValueError :
        temp_sheet_index = sheet_list.index( "temp" )
        sheet_obj = wb_obj.copy_worksheet ( wb_obj.worksheets[ temp_sheet_index ]  )
        sheet_obj.title = title_str
        wb_obj.save( file_name )
    else :
        sheet_obj = wb_obj.worksheets[ sheet_index ]
    finally:
        return sheet_obj

def ftp_service( ip, port, userName, password ):
    remote_file_path = diary_file
    ftp = FTP(host=ip , user=userName , passwd=password)
    ftp.login( user=userName , passwd=password)

    # 用户名登陆，OK
    # ftp = FTP(host="192.168.1.105",user="user",passwd="123")
    # ftp.login( user="user",passwd="123" )

    # 匿名登陆，文件写入 ftp-anonymous/ 下，OK
    # ftp = FTP(host="192.168.1.105")
    # ftp.set_debuglevel(2)
    # ftp.login()

    with open( file_name,"rb") as local_file :
        ftp.storbinary('STOR ' + remote_file_path ,local_file ,1024 )
    ftp.quit()

#################
def main():
    global file_name
    global cur_path

    cur_path  = sys.path[0]
    file_name  = cur_path+'/'+diary_file

    # 获取当前时间，包括年周数（一年内的第几周）和周日数（一周内的第几天）
    date_str,time_str,week_int,weekDay_int = get_week_date()
    # 根据年周数，确定表格名称
    sheet_title_str = str( week_int )

    # 提取命令行参数中的日志写入内容
    content_str = parser_content( sys.argv )

    ## 获取服务器配置
    auto_upload_int,server_ip_str,server_port_str,server_user_str,server_pass_str = parser_config() 


    # 获取工作簿及其中的对应表
    wb = get_workbook(  diary_file )
    sheet = get_sheet( wb , sheet_title_str )
    
    # 添加内容 
    privous_str = sheet.cell( weekDay_int+1 ,2).value
    # 检查是否是新cell，没有内容时，不追加新的内容
    if privous_str == None :
        sheet.cell( weekDay_int+1 ,2).value = "[ "+ time_str+"]"+ content_str
    else :
        sheet.cell( weekDay_int+1 ,2).value = privous_str +"\n[ "+ time_str+"]"+ content_str
    
    # 设置对齐格式
    sheet.cell( weekDay_int+1 ,2).alignment = Alignment( horizontal="left",vertical="top")
    wb.save(file_name)

    ## 检查自动上传的条件，在这一天之内，每次记录都会触发一次上传
    ## TODO：这个部分需要再优化 
    if  auto_upload_int == weekDay_int  :
        print ( " auto_upload_int == weekDay_int ,%s,%s,%s,%s" % ( server_ip_str, server_port_str, server_user_str, server_pass_str ) )
        ftp_service( server_ip_str, server_port_str, server_user_str, server_pass_str )
    return 

#################
## <<  程序入口 >> ##
if __name__ == "__main__":
    main()