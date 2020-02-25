简介  
本程序用于统计git下项目的代码提交量, 只适合于python3。

依懒库
pandas

运行  
python3 traverse_branch.py local_code_dir log  
local_code_dir: 本地clone下来的代码位置
log:统计好的数据存放的路径,每个分支统计结果单独保存到了一个文件

注意  
默认时间段: 2019年7月1日到当前。
分支, 当前本地的代码分支。
changeid, 有些项目好像是没有changeid的。