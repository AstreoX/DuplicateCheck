@echo off
echo 正在安装题库查重工具所需的依赖包...
echo.

REM 检查Python是否已安装
python --version 2>NUL
if %ERRORLEVEL% NEQ 0 (
    echo 错误: 未检测到Python安装，请先安装Python 3.7或更高版本。
    echo 您可以从 https://www.python.org/downloads/ 下载Python。
    pause
    exit /b 1
)

echo 检测到Python已安装，正在安装依赖...
echo.

REM 创建缓存目录
if not exist "%USERPROFILE%\.cache\sentence-transformers" (
    echo 创建模型缓存目录...
    mkdir "%USERPROFILE%\.cache\sentence-transformers"
)

REM 安装所需的Python库
echo 正在安装pandas...
pip install pandas
echo.

echo 正在安装numpy...
pip install numpy
echo.

echo 正在安装matplotlib...
pip install matplotlib
echo.

echo 正在安装tqdm...
pip install tqdm
echo.

echo 正在安装scikit-learn...
pip install scikit-learn
echo.

echo 正在安装xlrd (Excel读取支持)...
pip install xlrd
echo.

echo 正在安装openpyxl (Excel读取支持)...
pip install openpyxl
echo.

echo 正在安装sentence-transformers (这可能需要几分钟时间)...
pip install sentence-transformers
echo.

echo 所有依赖安装完成！
echo.
echo 您现在可以运行题库查重工具了。
echo.

pause