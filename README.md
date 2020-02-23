# xlsx 日志
用于随时记录工作内容，通过 xlsx 文件记录，以周为单位进行组织。
* 本工具基于“周”进行日记内容的组织；
* 每周的第一次记录，根据当前系统时间，自动建立一张新 sheet ；
* 一天中可以随时添加记录，通过当前时间查找对应的 sheet和行，记录内容将追加到当日记录的结尾。
 
 > Python 的开发确实快，边学边干，大半天就作出一个基本demo，后续逐步完善
 

## 依赖
* python3
* openpyxl

### 安装 pip
> 首先通过 wget 下载安装程序 get-pip.py \
$ wget https://bootstrap.pypa.io/get-pip.py \
以 python3 执行安装，自动获得 pip3 \
$ sudo python3 get-pip.py \
更新一下 pip \
$ pip install -U pip

### 安装 openpyxl
>通过 pip 安装 openpyxl \
$ pip install openpyxl

## 使用方法
首次启动该程序时，会自动在程序所处的路径下生成两个文件：\
* 运行配置文件 config.ini \
* 日志记录文件 your_file_name.xlsx

其中，配置文件需要填写部分内容，配置文件可手动修改该文件。
```sh
[DEFAULT]
auto_upload = 1                                     #自动推送日，设置为1～7，每周固定一天推送服务器
local_file_name = yangj_log.xlsx    # 推送文件，也就是日志内容文件

[FTP Server]
online = y                                  #在线模式，当没有服务器连接时，可以设为N，禁止自动推送
serverip = 192.168.1.105    #服务器地址
serverport = 21                         #服务器端口
username = yangj                #登陆FTP服务器的用户名
password = 111                      #用户密码
```

### 记录内容
通过命令参数 commit 提交一次记录，如果是第一次运行，则需要填写配置参数文件，并生成记录文件。
```sh
$ python3 diary.py commit
 首次使用,配置服务器参数：
 服务器IP : 12.12.12.12
 服务器Port : 12
 网络支持 (y/n): n
 用户名: yj
 用户密码: 1
 日志文件(*.xlsx): yj.xlsx
```
填写完成后，立即可以启动记录内容：
```sh
随手记]我的第一次记录内容就是这一句话
2020-02-23      第8周 第7天
```

### 查看内容
可通过 show 命令参数查看本周或指定周的日志内容，如下：
```sh
$ python3 diary.py show
 < 2020 年  第 8 周 > 
----------------------
|  周日   2020-03-01 |
----------------------
[21:07:38] 我的第一次记录内容就是这一句话

```
### 远程推送
通过 push 命令参数，启动一次远程服务器的推送。若远程ftp服务器关闭，则返回提示。
```sh
$ python3 diary.py push
```


> 在 .bashrc 中添加“命令别名”，重启终端或
>`
>$ source .bashrc
>`
>之后即可通过“别名”启动包含参数的工具，直接进入提交模式。
>```sh
># .bashrc 中添加如下别名
>alias job='python3 <your_path>/diary.py commit' 

## 测试
*测试部分不影响一般使用，仅仅用于开发过程。* \
在代码中添加 "test_*" 开头的函数，并通过 asssert 断言检查函数的输入和返回值的情况，以此进行程序的自动测试。
1. 安装 pytest 测试框架: \
`pip3 install -U pytest`

2. 项目中添加单元测试代码，如下：

```py
import pytest

def get_workbook( fileName ):
    if not os.access( fileName , os.F_OK ):
        print (" File Not Found Error ")
        return None
        # exit()
    return openpyxl.load_workbook( fileName )

###
def test_get_workbook():
    assert get_workbook("dddd")==None
```
