@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ===== 题库查重工具打包选择菜单 =====
echo.
echo 请选择打包方式：
echo 1. 使用PyInstaller打包（推荐）
echo 2. 使用cx_Freeze打包
echo 3. 退出
echo.

set /p choice=请输入选项（1-3）: 

if "%choice%"=="1" (
    echo.
    echo 您选择了使用PyInstaller打包
    echo 正在启动打包过程...
    python build_with_pyinstaller.py
) else if "%choice%"=="2" (
    echo.
    echo 您选择了使用cx_Freeze打包
    echo 正在启动打包过程...
    python setup.py build
) else if "%choice%"=="3" (
    echo.
    echo 退出程序
    exit /b 0
) else (
    echo.
    echo 无效的选项，请重新运行程序并选择正确的选项
    exit /b 1
)

if %errorlevel% neq 0 (
    echo.
    echo 打包过程中出现错误，请查看上方错误信息
) else (
    echo.
    echo 打包完成！可执行文件位于dist目录
)

echo.
echo 按任意键退出...
pause > nul