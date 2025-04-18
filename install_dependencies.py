"""
微信文件传输助手剪贴板监控工具 - 依赖安装程序

此脚本用于安装运行主程序所需的所有依赖库，
包括必需依赖和可选依赖。
"""

import sys
import subprocess
import importlib.util
import time
import os

print("=== 微信文件传输助手剪贴板监控工具 - 依赖安装程序 ===")
print("此脚本将安装程序运行所需的所有依赖库\n")

# 检查主程序文件是否存在
main_program = "wx_clipboard_monitor.py"
main_program_exists = os.path.exists(main_program)

if not main_program_exists:
    print(f"警告: 未找到主程序文件 {main_program}")
    print("请确保依赖安装程序与主程序在同一目录下\n")

# 必需的依赖包
required_packages = {
    'pyperclip': '剪贴板操作库',
    'pyautogui': '自动键盘鼠标操作库',
    'keyboard': '键盘监听和热键支持库'
}

# 可选但推荐的依赖包
optional_packages = {
    'pywin32': 'Windows API接口库，用于更精确地控制窗口'
}

# 检查和安装函数
def check_and_install_package(package_name, description):
    module_name = package_name
    if package_name == 'pywin32':
        module_name = 'win32gui'
    
    if importlib.util.find_spec(module_name) is None:
        print(f"- 正在安装 {package_name}({description})...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"  ✓ {package_name} 安装成功！")
            return True
        except Exception as e:
            print(f"  ✗ {package_name} 安装失败: {e}")
            return False
    else:
        print(f"- {package_name}({description}) 已安装")
        return True

# 安装必需依赖
print("\n== 安装必需依赖 ==")
required_success = True
for package, desc in required_packages.items():
    if not check_and_install_package(package, desc):
        required_success = False

# 安装可选依赖
print("\n== 安装可选依赖 ==")
optional_success = True
for package, desc in optional_packages.items():
    if not check_and_install_package(package, desc):
        optional_success = False

# 总结
print("\n== 安装结果 ==")
if required_success:
    print("✓ 所有必需依赖已成功安装！")
else:
    print("✗ 部分必需依赖安装失败，请查看上方错误信息")

if optional_success:
    print("✓ 所有可选依赖已成功安装！")
else:
    print("⚠ 部分可选依赖安装失败，但这不会影响基本功能")

print("\n如果安装过程中出现错误，您可以尝试手动安装依赖:")
for package in required_packages:
    print(f"  pip install {package}")
for package in optional_packages:
    print(f"  pip install {package}")

# 如果主程序存在，询问是否立即运行
if main_program_exists and required_success:
    print("\n所有必需依赖已安装完成。")
    
    try:
        choice = input("\n是否立即运行主程序？(y/n): ").strip().lower()
        if choice == 'y' or choice == 'yes':
            print("\n正在启动主程序...\n")
            subprocess.Popen([sys.executable, main_program])
            print("主程序已在新窗口中启动")
            time.sleep(1)
            sys.exit(0)
    except:
        pass

print("\n安装完成，5秒后将自动关闭...")
for i in range(5, 0, -1):
    print(f"\r倒计时: {i}秒", end="")
    time.sleep(1) 