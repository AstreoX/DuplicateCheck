import sys
import os
from cx_Freeze import setup, Executable

# 依赖包
dependencies = [
    "pandas",
    "numpy",
    "matplotlib",
    "tqdm",
    "sentence_transformers",
    "scikit_learn",
    "tkinter",
    "xlrd",
    "openpyxl"
]

# 构建选项
build_options = {
    "packages": dependencies,
    "excludes": [],
    "include_files": []
}

# 可执行文件
executables = [
    Executable(
        "question_duplicate_checker.py",  # 主脚本
        base="Win32GUI" if sys.platform == "win32" else None,  # 使用GUI模式
        target_name="题库查重工具.exe",  # 输出文件名
        icon=None,  # 图标文件
    )
]

# 设置
setup(
    name="题库查重工具",
    version="1.0",
    description="题库查重工具 - 自动检测题库中的重复题目",
    options={"build_exe": build_options},
    executables=executables
)