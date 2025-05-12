@echo off
echo ===== 题库查重工具打包脚本 =====
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python安装，请先安装Python
    pause
    exit /b 1
)

echo 开始打包题库查重工具为可执行文件...
echo 这可能需要几分钟时间，请耐心等待...
echo.

REM 运行打包脚本
python build_exe.py

if %errorlevel% neq 0 (
    echo 打包过程中出现错误，请查看上方错误信息
) else (
    echo.
    echo 打包完成！可执行文件位于dist目录
)

echo.
echo 按任意键退出...
pause > nul