import os
import sys
import subprocess
import time

def check_dependencies():
    """检查是否安装了所有必要的依赖"""
    required_packages = [
        "pandas",
        "numpy",
        "matplotlib",
        "tqdm",
        "sentence-transformers",
        "scikit-learn",
        "xlrd",
        "openpyxl",
        "pyinstaller"
    ]
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"缺少以下依赖包: {', '.join(missing_packages)}")
        install = input("是否自动安装这些依赖? (y/n): ")
        if install.lower() == 'y':
            for package in missing_packages:
                print(f"正在安装 {package}...")
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                    print(f"{package} 安装成功")
                except subprocess.CalledProcessError:
                    print(f"{package} 安装失败，请手动安装")
                    return False
            return True
        else:
            print("请先安装所有依赖后再运行此脚本")
            return False
    return True

def build_executable():
    """使用PyInstaller打包应用程序"""
    print("\n开始打包题库查重工具...")
    print("这可能需要几分钟时间，请耐心等待...\n")
    
    # 创建spec文件
    spec_content = f'''\
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['question_duplicate_checker.py'],
    pathex=[r'{os.getcwd()}'],
    binaries=[],
    datas=[('面向对象程序设计-题库 (1).xls', '.')],
    hiddenimports=['sentence_transformers', 'sklearn.metrics.pairwise'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='题库查重工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
'''
    
    with open('题库查重工具.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    # 使用PyInstaller打包
    try:
        subprocess.check_call([sys.executable, "-m", "PyInstaller", "题库查重工具.spec", "--clean"])
        print("\n打包完成！")
        exe_path = os.path.abspath(os.path.join('dist', '题库查重工具.exe'))
        print(f"可执行文件位于: {exe_path}")
        
        # 检查文件是否存在
        if os.path.exists(exe_path):
            print("\n打包成功！您可以双击运行该可执行文件。")
            return True
        else:
            print("\n警告：打包过程似乎完成，但未找到可执行文件。")
            return False
    except subprocess.CalledProcessError as e:
        print(f"\n打包失败: {e}")
        return False

def main():
    print("===== 题库查重工具打包程序 =====\n")
    
    # 检查依赖
    if not check_dependencies():
        print("\n依赖检查失败，打包过程终止")
        input("按Enter键退出...")
        return
    
    # 打包可执行文件
    start_time = time.time()
    success = build_executable()
    end_time = time.time()
    
    if success:
        print(f"\n打包用时: {end_time - start_time:.2f} 秒")
        print("\n您可以在dist目录找到打包好的可执行文件：题库查重工具.exe")
    else:
        print("\n打包过程失败，请检查错误信息并重试")
    
    input("\n按Enter键退出...")

if __name__ == "__main__":
    main()