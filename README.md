# xlsx 日志
用于随时记录工作内容，通过 xlsx 文件记录，以周为单位进行组织。

* 本工具基于“周”进行日记内容的组织；
* 每周的第一次记录，根据当前系统时间，自动建立一张新 sheet ；
* 一天中可以随时添加记录，通过当前时间查找对应的 sheet和行，记录内容将追加到当日记录的结尾。



## 依赖
* python3
本工具基于 python,需要首先安装 python3
* openpyxl
推荐通过 pip 安装 openpyxl，如下： 
```sh 
$ pip install openpyxl 
```

## 使用方法

0.  新建一个 log.xlsx 文件用于保存日志，该文件中必须包含一个名为 temp 的 sheet，内容如下：

|   日期    |  null  |
|---|---|
| 周一  | null  |
| 周二  | null  |
| 周三  | null  |
| 周四  | null  |
| 周五  | null  |
| 周六  | null  |
| 周日  | null  |

1. 将需要日志 log.xlsx 文件与本程序置于同一路径，执行以下命令即可：
```sh
$ python3 diary.py  commit
```

**强烈推荐**
> 在 .bashrc 中添加类似的“命令别名”，重启终端后执行 “别名”即可启动工具（包含参数），并直接进入提交模式。
```sh
# .bashrc 中添加如下别名
alias job='python3 <your_path>/diary.py commit' 

# 命令行执行即可
$ job
2020-02-03  20:20:32  6th WEEK 1th Day 
记录工作内容：
```

## 不得不说
Python 的开发确实快，边学边干，半天就作出这个demo（后续还有很多可以作的）
