chcp 65001
@echo off
echo "请注意必须使用管理员运行该命令文件，且只需要执行一次即可！"
echo "程序安装中…… >>>"
set current_path=%~dp0
echo "开始先卸载已存在的注册" || echo "卸载注册程序失败"
%current_path%CWeChatRobot.exe /unregserver
echo "开始注册程序" || echo "注册程序失败"
%current_path%CWeChatRobot.exe /regserver
echo "开始安装依赖"
pip3 install -r %current_path%requirements.txt -i https://pypi.douban.com/simple || (
    echo "安装依赖失败，尝试备用源安装"
    pip3 install -r %current_path%requirements.txt
)
echo "全部操作执行完毕，可以关闭本窗口了！"
pause
