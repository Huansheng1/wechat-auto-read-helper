chcp 65001
@echo off
echo "程序初始化中……正在开机……欢迎使用 幻生公开版自动过检测机器人 >>>"
set current_path=%~dp0
echo "开始启动程序界面……"
REM 检测Python 3是否存在
where python3 >nul 2>nul
if %errorlevel% equ 0 (
    python3 %current_path%\wechat_bot_gui.py
) else (
    python %current_path%\wechat_bot_gui.py
)
echo "全部操作执行完毕，不可关闭本窗口！"
pause