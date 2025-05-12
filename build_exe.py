import os
import sys
import subprocess
import shutil

def check_pyinstaller():
    """检查是否安装了PyInstaller，如果没有则安装"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "show", "pyinstaller"], 
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("PyInstaller 已安装")
        return True
    except subprocess.CalledProcessError:
        print("正在安装 PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller 安装成功")
            return True
        except subprocess.CalledProcessError:
            print("PyInstaller 安装失败，请手动安装后重试")
            return False

def build_executable():
    """使用PyInstaller打包应用程序"""
    if not check_pyinstaller():
        return False
    
    print("\n开始打包应用程序...")
    print("这可能需要几分钟时间，请耐心等待...\n")
    
    # 创建输出目录
    if not os.path.exists("dist"):
        os.makedirs("dist")
    
    # 使用PyInstaller打包
    cmd = [
        sys.executable, "-m", "pyinstaller",
        "--name=题库查重工具",
        "--windowed",  # 使用GUI模式
        "--onefile",   # 打包成单个文件
        "--clean",     # 清理临时文件
        "--add-data=面向对象程序设计-题库 (1).xls;.",  # 添加示例数据文件
        "question_duplicate_checker.py"  # 主脚本
    ]
    
    try:
        subprocess.check_call(cmd)
        print("\n打包完成！")
        print(f"可执行文件位于: {os.path.abspath('dist/题库查重工具.exe')}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        return False

def update_readme():
    """更新README文件，添加可执行文件使用说明"""
    readme_path = "README.md"
    
    if not os.path.exists(readme_path):
        print("未找到README.md文件，将创建新文件")
        with open(readme_path, "w", encoding="utf-8") as f:
            f.write("# 题库查重工具\n\n")
    
    # 读取现有内容
    with open(readme_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # 添加可执行文件使用说明
    exe_instructions = """
## 可执行文件使用说明

为方便不熟悉Python的用户，我们提供了打包好的可执行文件版本。

### 使用方法

1. 下载 `题库查重工具.exe` 文件
2. 双击运行该文件
3. 在界面中选择Excel题库文件
4. 设置相似度阈值
5. 点击"开始分析"按钮

### 注意事项

- 可执行文件包含所有必要的依赖，无需安装Python或其他库
- 首次运行时，Windows可能会显示安全警告，请选择"仍要运行"
- 如需使用源代码版本，请参考上方的Python安装说明

### 打包自己的版本

如果您想自行打包可执行文件，可以运行以下命令：

```
python build_exe.py
```

打包完成后，可执行文件将位于 `dist` 目录中。
"""
    
    # 检查是否已存在可执行文件说明
    if "## 可执行文件使用说明" not in content:
        with open(readme_path, "a", encoding="utf-8") as f:
            f.write(exe_instructions)
        print("已更新README.md，添加了可执行文件使用说明")
    else:
        print("README.md已包含可执行文件使用说明，无需更新")

def main():
    """主函数"""
    print("=== 题库查重工具打包脚本 ===")
    
    # 打包可执行文件
    if build_executable():
        # 更新README
        update_readme()
        print("\n打包过程已完成！")
        print("您可以在dist目录找到打包好的可执行文件：题库查重工具.exe")
    else:
        print("\n打包过程失败，请检查错误信息并重试")

if __name__ == "__main__":
    main()