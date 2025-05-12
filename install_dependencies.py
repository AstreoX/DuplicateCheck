import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

def main():
    # 要安装的包列表
    packages = [
        "pandas",
        "numpy",
        "matplotlib",
        "tqdm",
        "sentence-transformers",
        "scikit-learn",
        "xlrd",
        "openpyxl"
    ]

    # 显示安装信息
    print("开始安装依赖包...\n")

    # 安装包
    success_count = 0
    for package in packages:
        print(f"正在安装 {package}...")

        if install_package(package):
            print("成功!\n")
            success_count += 1
        else:
            print("失败!\n")

    # 显示安装结果
    if success_count == len(packages):
        print("\n所有依赖包安装成功!\n")
        print("\n现在您可以运行 question_duplicate_checker.py 来启动题库查重工具。\n")
    else:
        print(f"\n安装完成，但有 {len(packages) - success_count} 个包安装失败。\n")
        print("请手动安装失败的包，或者尝试以管理员身份运行此脚本。\n")

if __name__ == "__main__":
    main()