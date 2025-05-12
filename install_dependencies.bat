@echo off

REM 要安装的包列表
set packages=pandas numpy matplotlib tqdm sentence-transformers scikit-learn xlrd openpyxl

REM 显示安装信息
echo 开始安装依赖包...

REM 安装包
set success_count=0
for %%p in (%packages%) do (
    echo 正在安装 %%p...
    pip install %%p
    if errorlevel 1 (
        echo 失败!
    ) else (
        echo 成功!
        set /a success_count+=1
    )
)

REM 显示安装结果
if %success_count%==8 (
    echo.
    echo 所有依赖包安装成功!
    echo.
    echo 现在您可以运行 question_duplicate_checker.py 来启动题库查重工具。
) else (
    echo.
    echo 安装完成，但有 %success_count% 个包安装失败。
    echo 请手动安装失败的包，或者尝试以管理员身份运行此脚本。
)