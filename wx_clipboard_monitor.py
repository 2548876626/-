"""
微信文件传输助手剪贴板监控工具
自动监控剪贴板，当检测到学习验证链接时，提示用户手动切换到微信文件传输助手，然后按快捷键触发自动发送
"""

import sys
import subprocess
import importlib.util
import os
import json
from datetime import datetime, timedelta

# 检测是否是从命令行直接运行（而不是被导入）
is_main_run = __name__ == "__main__"

# 配置文件路径
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".wx_clipboard_monitor")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
DEPENDENCIES_CHECK_FILE = os.path.join(CONFIG_DIR, "dependencies_check.json")
FIRST_RUN_FLAG_FILE = os.path.join(CONFIG_DIR, "first_run_completed")

# 确保配置目录存在
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)

# 检测是否首次运行
is_first_run = not os.path.exists(FIRST_RUN_FLAG_FILE)

# 如果是第一次运行且从命令行启动，显示欢迎信息
if is_first_run and is_main_run:
    print("\n" + "="*70)
    print(" "*10 + "欢迎使用微信文件传输助手剪贴板监控工具" + " "*10)
    print("="*70)
    print("首次启动将自动检查并安装必要的依赖库，这可能需要一点时间...")
    print("\n如果您希望手动安装依赖，可以运行 install_dependencies.py 脚本\n")

# 检查必要的依赖
required_packages = {
    'pyperclip': 'pyperclip',
    'pyautogui': 'pyautogui',
    'keyboard': 'keyboard'
}

# 可选依赖（如果可用则使用，不可用则使用替代方案）
optional_packages = {
    'pywin32': 'win32gui'
}

def check_and_install_dependencies(force_check=False):
    """检查并安装必要的依赖库，带智能检测功能"""
    # 检查上次检测时间，如果在24小时内已检测过，且不是强制检测或首次运行，则跳过
    skip_check = False
    if not force_check and not is_first_run and os.path.exists(DEPENDENCIES_CHECK_FILE):
        try:
            with open(DEPENDENCIES_CHECK_FILE, 'r') as f:
                check_data = json.load(f)
                last_check = datetime.fromisoformat(check_data.get('last_check', ''))
                if datetime.now() - last_check < timedelta(hours=24):
                    if is_main_run:
                        print("距离上次依赖检查不到24小时，跳过检查")
                    skip_check = True
        except:
            pass
    
    if skip_check:
        return True
    
    missing_packages = []
    
    # 检查必需包
    for package_name, import_name in required_packages.items():
        if importlib.util.find_spec(import_name) is None:
            missing_packages.append(package_name)
    
    # 检查可选包（仅记录，不强制安装）
    optional_missing = []
    for package_name, import_name in optional_packages.items():
        if importlib.util.find_spec(import_name) is None:
            optional_missing.append(package_name)
    
    # 安装必需包
    if missing_packages:
        if is_main_run:
            print(f"\n检测到缺少以下必要依赖: {', '.join(missing_packages)}")
            print("开始自动安装...\n")
        
        try:
            for package in missing_packages:
                if is_main_run:
                    print(f"正在安装 {package}...")
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", package],
                    stdout=subprocess.DEVNULL if not is_first_run else None,
                    stderr=subprocess.DEVNULL if not is_first_run else None
                )
                if is_main_run:
                    print(f"✓ {package} 安装成功")
            
            if is_main_run:
                print("\n所有必需依赖已安装完成")
        except Exception as e:
            if is_main_run:
                print(f"\n安装依赖失败: {e}")
                print("\n请手动安装以下依赖后重新运行程序:")
                for package in missing_packages:
                    print(f"  pip install {package}")
                
                print("\n或者运行 install_dependencies.py 脚本安装所有依赖")
                
                if is_first_run:
                    input("\n按Enter键退出...")
            
            return False
    elif is_first_run and is_main_run:
        print("\n所有必需依赖已安装，无需额外安装")
    
    # 提示可选包（但不强制安装）
    if optional_missing and is_main_run and is_first_run:
        print(f"\n以下可选依赖未安装: {', '.join(optional_missing)}")
        print("这些库不是必须的，程序会使用替代方法，但安装它们可能提高兼容性")
        print("如需安装，请运行 install_dependencies.py 脚本或手动执行:")
        for package in optional_missing:
            print(f"  pip install {package}")
    
    # 记录此次检查时间
    try:
        with open(DEPENDENCIES_CHECK_FILE, 'w') as f:
            json.dump({
                'last_check': datetime.now().isoformat(),
                'missing_optional': optional_missing
            }, f)
    except:
        pass
    
    # 如果是首次运行，创建标志文件
    if is_first_run:
        try:
            with open(FIRST_RUN_FLAG_FILE, 'w') as f:
                f.write(datetime.now().isoformat())
        except:
            pass
    
    # 所有必需依赖已安装，继续运行
    if is_main_run and is_first_run:
        print("\n依赖检查完成，程序将继续运行")
        print("="*50 + "\n")
    
    return True

# 检查依赖
dependencies_ok = check_and_install_dependencies()
if not dependencies_ok:
    sys.exit(1)

# 导入必要的库
import pyperclip
import pyautogui
import time
import keyboard
import threading
import re
import tkinter as tk
from tkinter import messagebox

# 用户配置和窗口设置
USER_SETTINGS = {
    "target_url": "https://gd.aqscwlxy.com/h5/pages/pc/pcPhoto",  # 要监测的URL
    "check_interval": 1.0,  # 检测间隔，秒
    "toggle_hotkey": "ctrl+shift+m",  # 切换监控状态的热键
    "send_hotkey": "ctrl+alt+s",      # 发送到微信的热键 (按此键自动粘贴发送)
    "saved_windows": []               # 保存的窗口标题列表
}

# 加载用户设置
def load_user_settings():
    """从配置文件加载用户设置"""
    global USER_SETTINGS
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved_settings = json.load(f)
                for key, value in saved_settings.items():
                    USER_SETTINGS[key] = value
            print(f"已加载用户配置: {CONFIG_FILE}")
        except Exception as e:
            print(f"加载配置文件失败: {e}")

# 保存用户设置
def save_user_settings():
    """保存用户设置到配置文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(USER_SETTINGS, f, ensure_ascii=False, indent=2)
        print(f"已保存用户配置: {CONFIG_FILE}")
    except Exception as e:
        print(f"保存配置文件失败: {e}")

# 加载用户设置
load_user_settings()

# 配置参数
CONFIG = {
    "target_url": USER_SETTINGS["target_url"],
    "check_interval": USER_SETTINGS["check_interval"],
    "toggle_hotkey": USER_SETTINGS["toggle_hotkey"],
    "send_hotkey": USER_SETTINGS["send_hotkey"],
}

# 全局变量
last_clipboard_content = ""
last_processed_content = ""
is_monitoring = True
processed_url_ready = False  # 标记是否有处理好的链接等待发送
selected_wechat_window = None  # 存储用户选择的微信窗口
selected_window_title = ""     # 存储选中窗口的标题
root = None
status_label = None
status_indicator = None
log_text = None

def log_message(message):
    """记录日志消息"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_msg = f"[{timestamp}] {message}"
    print(log_msg)
    
    # 如果GUI已初始化，也更新GUI日志
    if log_text:
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, log_msg + "\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)

def process_text(text):
    """处理文本，不再删除HTML转义字符"""
    # 不再进行任何处理，直接返回原始文本
    log_message("检测到剪贴板内容变化，无需处理HTML转义字符")
    return text

def send_message():
    """发送当前处理好的消息"""
    global processed_url_ready, last_processed_content, selected_wechat_window
    
    if not processed_url_ready or not last_processed_content:
        log_message("没有待发送的链接")
        show_notification("没有待发送的链接", "warning")
        return False
    
    try:
        log_message(f"准备发送文本: {last_processed_content[:50]}..." if len(last_processed_content) > 50 else f"准备发送文本: {last_processed_content}")
        
        # 确保最新处理的内容在剪贴板中
        pyperclip.copy(last_processed_content)
        time.sleep(0.5)  # 增加延迟
        
        # 尝试多种方法发送消息
        method_success = False
        
        # 方法0：如果用户已选择窗口，优先使用该窗口（使用ctypes实现，不依赖win32gui）
        if selected_wechat_window:
            try:
                log_message(f"方法0: 使用用户选择的窗口 (hwnd: {selected_wechat_window})")
                
                # 使用ctypes激活窗口
                try:
                    import ctypes
                    user32 = ctypes.windll.user32
                    
                    # 尝试获取窗口标题
                    try:
                        title_length = user32.GetWindowTextLengthW(selected_wechat_window) + 1
                        title_buffer = ctypes.create_unicode_buffer(title_length)
                        user32.GetWindowTextW(selected_wechat_window, title_buffer, title_length)
                        window_title = title_buffer.value
                        log_message(f"已找到选择的窗口: {window_title}")
                    except:
                        log_message("无法获取窗口标题，但将继续尝试激活窗口")
                    
                    # 激活窗口
                    SW_RESTORE = 9  # 恢复窗口
                    user32.ShowWindow(selected_wechat_window, SW_RESTORE)
                    user32.SetForegroundWindow(selected_wechat_window)
                    time.sleep(0.5)
                    
                    # 粘贴并发送
                    log_message("执行粘贴操作")
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.7)
                    
                    log_message("执行发送操作")
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    
                    method_success = True
                    log_message("方法0成功：通过用户选择的窗口发送消息")
                except Exception as e:
                    log_message(f"使用ctypes激活窗口失败: {e}")
                    
                    # 尝试使用win32gui作为备选方案
                    try:
                        import win32gui
                        import win32con
                        
                        log_message("尝试使用win32gui激活窗口")
                        win32gui.ShowWindow(selected_wechat_window, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(selected_wechat_window)
                        time.sleep(0.5)
                        
                        # 粘贴并发送
                        log_message("执行粘贴操作")
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.7)
                        
                        log_message("执行发送操作")
                        pyautogui.press('enter')
                        time.sleep(0.5)
                        
                        method_success = True
                        log_message("方法0成功：通过win32gui激活窗口发送消息")
                    except Exception as e:
                        log_message(f"使用win32gui激活窗口也失败了: {e}")
            except Exception as e:
                log_message(f"方法0失败: {e}")
        
        # 如果方法0失败，尝试方法1：使用win32gui查找微信窗口
        if not method_success:
            try:
                log_message("方法1: 尝试查找并激活微信窗口")
                import win32gui
                import win32con
                
                # 尝试查找微信窗口
                def find_wechat_window(hwnd, results):
                    window_text = win32gui.GetWindowText(hwnd)
                    if '微信' in window_text or 'WeChat' in window_text or '文件传输助手' in window_text:
                        log_message(f"找到可能的微信窗口: {window_text}")
                        results.append(hwnd)
                    return True
                
                wechat_windows = []
                win32gui.EnumWindows(find_wechat_window, wechat_windows)
                
                if wechat_windows:
                    for hwnd in wechat_windows:
                        try:
                            log_message(f"尝试激活窗口: {win32gui.GetWindowText(hwnd)}")
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.5)
                            
                            # 粘贴并发送
                            log_message("执行粘贴操作")
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(0.7)
                            
                            log_message("执行发送操作")
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            
                            method_success = True
                            log_message("方法1成功：通过激活微信窗口发送消息")
                            break
                        except Exception as e:
                            log_message(f"激活窗口失败: {e}")
                            continue
                else:
                    log_message("未找到微信窗口")
            except Exception as e:
                log_message(f"方法1失败: {e}")
        
        # 方法2：如果前面的方法失败，尝试使用pyautogui模拟Alt+Tab操作
        if not method_success:
            try:
                log_message("方法2: 尝试使用Alt+Tab切换窗口")
                # 模拟Alt+Tab切换到之前的窗口，希望是微信
                pyautogui.keyDown('alt')
                pyautogui.press('tab')
                pyautogui.keyUp('alt')
                time.sleep(0.5)
                
                # 确认当前剪贴板内容
                current_clip = pyperclip.paste()
                if current_clip != last_processed_content:
                    log_message("警告：剪贴板内容可能已被更改，重新复制")
                    pyperclip.copy(last_processed_content)
                    time.sleep(0.3)
                
                # 粘贴并发送
                log_message("执行粘贴操作")
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.7)
                
                log_message("执行发送操作")
                pyautogui.press('enter')
                time.sleep(0.5)
                
                method_success = True
                log_message("方法2成功：通过Alt+Tab切换窗口发送消息")
            except Exception as e:
                log_message(f"方法2失败: {e}")
        
        # 方法3：如果前面的方法都失败，尝试打开微信文件传输助手网页版
        if not method_success:
            try:
                log_message("方法3: 尝试打开微信文件传输助手网页版")
                import webbrowser
                
                # 打开微信文件传输助手网页版
                webbrowser.open("https://filehelper.weixin.qq.com/")
                time.sleep(3)  # 等待网页加载
                
                # 尝试定位输入框并粘贴发送
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.7)
                
                # 查找并点击发送按钮
                # 由于网页版界面可能会变化，这里使用Enter键尝试发送
                pyautogui.press('enter')
                time.sleep(0.5)
                
                method_success = True
                log_message("方法3成功：通过网页版文件传输助手发送消息")
            except Exception as e:
                log_message(f"方法3失败: {e}")
        
        if method_success:
            log_message("消息已成功发送")
            show_notification("链接已成功发送到微信", "success")
            processed_url_ready = False
            play_alert_sound()
            return True
        else:
            log_message("所有发送方法都失败，请手动发送")
            show_notification("自动发送失败，请手动将剪贴板内容发送到微信", "error")
            return False
            
    except Exception as e:
        log_message(f"发送消息出错: {e}")
        show_notification(f"发送失败: {e}", "error")
        return False

def check_clipboard():
    """检查剪贴板内容"""
    global last_clipboard_content, last_processed_content, processed_url_ready
    
    if not is_monitoring:
        return
    
    try:
        # 读取剪贴板内容
        text = pyperclip.paste()
        
        # 如果内容为空或与上次相同，不处理
        if not text or text == last_clipboard_content:
            return
        
        # 更新上次检测到的内容
        last_clipboard_content = text
        
        # 处理文本
        processed_text = process_text(text)
        
        # 检查处理后的文本是否包含目标URL
        if CONFIG["target_url"] in processed_text:
            log_message("检测到目标URL!")
            
            # 如果和上次处理过的内容不同，才进行处理
            if processed_text != last_processed_content:
                # 记录这次处理的内容
                last_processed_content = processed_text
                processed_url_ready = True
                
                # 自动复制处理后的内容到剪贴板
                pyperclip.copy(processed_text)
                
                # 显示通知
                show_notification(
                    f"检测到学习验证链接！\n请切换到微信文件传输助手，然后按 {CONFIG['send_hotkey']} 发送", 
                    "success"
                )
                play_alert_sound()
            else:
                log_message("该链接已处理过，跳过")
    except Exception as e:
        log_message(f"检查剪贴板时出错: {e}")

def toggle_monitoring():
    """切换监控状态"""
    global is_monitoring
    is_monitoring = not is_monitoring
    update_status_indicator()
    
    show_notification(
        "剪贴板监控已启动" if is_monitoring else "剪贴板监控已暂停",
        "success" if is_monitoring else "warning"
    )
    
    log_message("监测已启动" if is_monitoring else "监测已暂停")

def play_alert_sound():
    """播放提示音"""
    try:
        import winsound
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
    except Exception as e:
        log_message(f"播放提示音失败: {e}")

def show_notification(message, type="info"):
    """显示通知"""
    log_message(message)
    
    # 如果GUI已初始化，使用GUI显示通知
    if root:
        # 创建通知颜色
        if type == "success":
            bg_color = "#4CAF50"
        elif type == "error":
            bg_color = "#F44336"
        elif type == "warning":
            bg_color = "#FF9800"
        else:
            bg_color = "#2196F3"
        
        # 创建通知窗口
        notification = tk.Toplevel(root)
        notification.overrideredirect(True)
        notification.attributes('-topmost', True)
        
        # 设置通知位置（屏幕右下角）
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        notification.geometry(f"300x100+{screen_width-320}+{screen_height-120}")
        
        # 设置通知内容
        notification.configure(bg=bg_color)
        
        # 消息标签
        message_label = tk.Label(
            notification, 
            text=message, 
            bg=bg_color, 
            fg="white", 
            font=("Arial", 12, "bold"),
            wraplength=280,
            padx=10, 
            pady=10
        )
        message_label.pack(fill=tk.BOTH, expand=True)
        
        # 如果是检测到链接的通知，添加发送按钮
        if "检测到学习验证链接" in message:
            send_button = tk.Button(
                notification,
                text="立即发送",
                bg="#FFFFFF",
                fg=bg_color,
                relief=tk.FLAT,
                padx=10,
                pady=5,
                font=("Arial", 10, "bold"),
                command=lambda: [send_message(), notification.destroy()]
            )
            send_button.pack(pady=(0, 10))
        
        # 5秒后关闭
        notification.after(5000, notification.destroy)
    else:
        # 如果GUI未初始化，使用messagebox
        if type == "error":
            messagebox.showerror("错误", message)
        elif type == "warning":
            messagebox.showwarning("警告", message)
        else:
            messagebox.showinfo("提示", message)

def update_status_indicator():
    """更新状态指示器"""
    if status_indicator and status_label:
        status_indicator.configure(bg="#4CAF50" if is_monitoring else "#F44336")
        status_label.configure(text="监控已启动" if is_monitoring else "监控已暂停")

def create_gui():
    """创建GUI界面"""
    global root, status_label, status_indicator, log_text
    
    # 创建主窗口
    root = tk.Tk()
    root.title("微信文件传输助手剪贴板监控")
    root.geometry("500x400")
    root.resizable(True, True)
    
    # 创建顶部状态栏
    status_frame = tk.Frame(root, padx=10, pady=5)
    status_frame.pack(fill=tk.X)
    
    # 状态指示器
    status_indicator = tk.Canvas(status_frame, width=12, height=12, bg="#4CAF50")
    status_indicator.pack(side=tk.LEFT, padx=(0, 5))
    
    # 绘制圆形
    status_indicator.create_oval(2, 2, 10, 10, fill="#4CAF50", outline="")
    
    # 状态文本
    status_label = tk.Label(status_frame, text="监控已启动")
    status_label.pack(side=tk.LEFT)
    
    # 分隔线
    tk.Frame(root, height=1, bg="#E0E0E0").pack(fill=tk.X, padx=10, pady=5)
    
    # 监控目标信息
    info_frame = tk.Frame(root, padx=10, pady=5)
    info_frame.pack(fill=tk.X)
    
    tk.Label(info_frame, text=f"监控URL: {CONFIG['target_url']}").pack(anchor=tk.W)
    tk.Label(info_frame, text=f"开/关监控: {CONFIG['toggle_hotkey']}").pack(anchor=tk.W)
    tk.Label(info_frame, text=f"发送快捷键: {CONFIG['send_hotkey']}").pack(anchor=tk.W)
    
    # 分隔线
    tk.Frame(root, height=1, bg="#E0E0E0").pack(fill=tk.X, padx=10, pady=5)
    
    # 控制按钮 - 第一行
    button_frame1 = tk.Frame(root, padx=10, pady=5)
    button_frame1.pack(fill=tk.X)
    
    toggle_button = tk.Button(
        button_frame1, 
        text="暂停监控" if is_monitoring else "启动监控", 
        command=toggle_monitoring,
        bg="#2196F3",
        fg="white",
        padx=10
    )
    toggle_button.pack(side=tk.LEFT, padx=(0, 10))
    
    # 发送按钮 - 仅当有已处理的链接时可用
    send_button = tk.Button(
        button_frame1,
        text="发送到微信",
        command=send_message,
        bg="#4CAF50",
        fg="white",
        padx=10
    )
    send_button.pack(side=tk.LEFT, padx=(0, 10))
    
    # 清空日志按钮
    clear_log_button = tk.Button(
        button_frame1,
        text="清空日志",
        command=lambda: log_text.delete(1.0, tk.END) if log_text else None,
        padx=10
    )
    clear_log_button.pack(side=tk.LEFT)
    
    # 窗口选择按钮 - 第二行
    button_frame2 = tk.Frame(root, padx=10, pady=0)
    button_frame2.pack(fill=tk.X)
    
    # 手动选择窗口按钮
    manual_select_button = tk.Button(
        button_frame2,
        text="手动选择窗口(推荐)",
        command=manual_select_window,
        bg="#FF9800",
        fg="white",
        padx=10,
        font=("Arial", 9, "bold")
    )
    manual_select_button.pack(side=tk.LEFT, padx=(0, 10), pady=5)
    
    # 自动选择窗口按钮
    auto_select_button = tk.Button(
        button_frame2,
        text="自动选择窗口",
        command=select_wechat_window,
        bg="#FF9800",
        fg="white",
        padx=10
    )
    auto_select_button.pack(side=tk.LEFT, padx=(0, 10), pady=5)
    
    # 日志区域
    log_frame = tk.Frame(root, padx=10, pady=5)
    log_frame.pack(fill=tk.BOTH, expand=True)
    
    tk.Label(log_frame, text="操作日志:").pack(anchor=tk.W)
    
    log_text = tk.Text(log_frame, height=10, width=50, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True)
    
    # 添加滚动条
    scrollbar = tk.Scrollbar(log_text)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    log_text.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=log_text.yview)
    
    # 底部状态栏
    footer_frame = tk.Frame(root, padx=10, pady=5, bg="#F5F5F5")
    footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
    
    # 添加窗口状态信息
    window_status_var = tk.StringVar(value="未选择窗口")
    window_status_label = tk.Label(
        footer_frame, 
        textvariable=window_status_var,
        fg="#757575", 
        bg="#F5F5F5"
    )
    window_status_label.pack(side=tk.RIGHT)
    
    # 更新窗口状态显示
    def update_window_status():
        global selected_wechat_window
        if selected_wechat_window:
            try:
                # 尝试使用ctypes获取窗口标题
                import ctypes
                user32 = ctypes.windll.user32
                
                # 获取窗口标题的长度
                title_length = user32.GetWindowTextLengthW(selected_wechat_window) + 1
                title_buffer = ctypes.create_unicode_buffer(title_length)
                user32.GetWindowTextW(selected_wechat_window, title_buffer, title_length)
                window_title = title_buffer.value
                
                window_status_var.set(f"已选择窗口: {window_title[:20]}..." if len(window_title) > 20 else f"已选择窗口: {window_title}")
            except:
                try:
                    # 备用方法：使用win32gui
                    import win32gui
                    window_title = win32gui.GetWindowText(selected_wechat_window)
                    window_status_var.set(f"已选择窗口: {window_title[:20]}..." if len(window_title) > 20 else f"已选择窗口: {window_title}")
                except:
                    window_status_var.set("已选择窗口(未知标题)")
        else:
            window_status_var.set("未选择窗口")
        
        root.after(1000, update_window_status)
    
    update_window_status()
    
    tk.Label(
        footer_frame, 
        text="微信文件传输助手剪贴板监控 v1.0", 
        fg="#757575",
        bg="#F5F5F5"
    ).pack(side=tk.LEFT)
    
    # 保持窗口响应
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 更新后续控制逻辑
    def update_button_state():
        send_button.config(state=tk.NORMAL if processed_url_ready else tk.DISABLED)
        root.after(1000, update_button_state)
    
    update_button_state()
    
    return root

def on_closing():
    """窗口关闭事件处理"""
    if messagebox.askokcancel("确认", "是否关闭监控程序?"):
        root.destroy()
        # 停止线程
        if monitor_thread and monitor_thread.is_alive():
            global is_monitoring
            is_monitoring = False  # 这会让线程循环退出

def monitor_clipboard_thread():
    """剪贴板监控线程"""
    while True:
        try:
            if not is_monitoring:
                time.sleep(1)
                continue
                
            check_clipboard()
            time.sleep(CONFIG["check_interval"])
        except Exception as e:
            log_message(f"监控线程出错: {e}")
            time.sleep(CONFIG["check_interval"])

def main():
    """主函数"""
    global monitor_thread
    
    # 注册热键
    keyboard.add_hotkey(CONFIG["toggle_hotkey"], toggle_monitoring)
    keyboard.add_hotkey(CONFIG["send_hotkey"], send_message)
    
    # 创建GUI
    gui = create_gui()
    
    # 尝试恢复上次选择的窗口
    if restore_saved_window():
        log_message(f"已自动恢复上次选择的窗口: {selected_window_title}")
    
    # 启动监控线程
    monitor_thread = threading.Thread(target=monitor_clipboard_thread, daemon=True)
    monitor_thread.start()
    
    # 显示初始通知
    show_notification("剪贴板监控已启动", "success")
    log_message("=== 微信文件传输助手剪贴板监控工具已启动 ===")
    log_message(f"正在监测URL: {CONFIG['target_url']}")
    log_message(f"按 {CONFIG['toggle_hotkey']} 可切换监测状态")
    log_message(f"按 {CONFIG['send_hotkey']} 可发送已处理的链接")
    log_message("=== 使用说明 ===")
    log_message("1. 当检测到学习验证链接时，会自动提示")
    log_message("2. 点击「手动选择窗口」按钮，然后点击微信或文件传输助手窗口")
    log_message("3. 检测到链接后，点击「发送到微信」按钮或按快捷键发送")
    log_message("注意：选择的窗口会自动保存，下次启动时自动恢复")
    log_message("======================================")
    
    # 启动GUI主循环
    gui.mainloop()

# 新增函数：列出所有窗口并让用户选择微信窗口
def select_wechat_window():
    """列出所有可能的微信窗口并让用户选择"""
    global selected_wechat_window
    
    try:
        import win32gui
        
        # 获取所有窗口
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if window_text and len(window_text.strip()) > 0:
                    windows.append((hwnd, window_text))
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        # 过滤可能的微信窗口
        wechat_windows = []
        other_windows = []
        
        for hwnd, title in windows:
            if '微信' in title or 'WeChat' in title or '文件传输助手' in title:
                wechat_windows.append((hwnd, title))
            elif title.strip():  # 其他非空标题窗口
                other_windows.append((hwnd, title))
        
        # 创建窗口选择对话框
        select_window = tk.Toplevel(root)
        select_window.title("选择微信窗口")
        select_window.geometry("500x400")
        select_window.grab_set()  # 模态对话框
        
        tk.Label(select_window, text="请选择微信窗口或文件传输助手窗口:", font=("Arial", 12)).pack(pady=10)
        
        # 创建列表框
        listbox_frame = tk.Frame(select_window)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(listbox_frame, font=("Arial", 10), yscrollcommand=scrollbar.set)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 首先添加微信窗口
        listbox.insert(tk.END, "--- 可能的微信窗口 ---")
        for idx, (hwnd, title) in enumerate(wechat_windows):
            listbox.insert(tk.END, f"{title} (hwnd: {hwnd})")
        
        # 然后添加其他窗口
        listbox.insert(tk.END, "--- 其他窗口 ---")
        for idx, (hwnd, title) in enumerate(other_windows):
            listbox.insert(tk.END, f"{title} (hwnd: {hwnd})")
        
        # 选中第一个微信窗口（如果有）
        if len(wechat_windows) > 0:
            listbox.selection_set(1)  # 跳过标题行，选择第一个微信窗口
        
        # 确认按钮处理函数
        def on_confirm():
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                if idx == 0 or idx == len(wechat_windows) + 1:  # 如果选中的是标题行
                    messagebox.showwarning("警告", "请选择一个有效的窗口")
                    return
                
                if idx <= len(wechat_windows):  # 选中的是微信窗口
                    selected_hwnd, title = wechat_windows[idx - 1]  # 减1是因为有标题行
                else:  # 选中的是其他窗口
                    other_idx = idx - len(wechat_windows) - 2  # 减2是因为有两个标题行
                    selected_hwnd, title = other_windows[other_idx]
                
                selected_wechat_window = selected_hwnd
                log_message(f"已选择窗口: {title} (hwnd: {selected_hwnd})")
                show_notification(f"已选择窗口: {title}", "success")
                select_window.destroy()
            else:
                messagebox.showwarning("警告", "请选择一个窗口")
        
        # 取消按钮处理函数
        def on_cancel():
            select_window.destroy()
        
        # 添加按钮
        button_frame = tk.Frame(select_window)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="确认", command=on_confirm, bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=on_cancel, bg="#F44336", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        
        # 窗口关闭事件
        select_window.protocol("WM_DELETE_WINDOW", on_cancel)
        
        # 两击选择
        def on_double_click(event):
            on_confirm()
        
        listbox.bind("<Double-Button-1>", on_double_click)
        
        return True
    except Exception as e:
        log_message(f"无法列出窗口: {e}")
        show_notification(f"无法列出窗口: {e}", "error")
        return False

# 新增函数：手动选择窗口（无需win32gui）
def manual_select_window():
    """手动选择窗口，不依赖win32gui"""
    global selected_wechat_window, selected_window_title
    
    # 创建一个简单的指示窗口
    guide_window = tk.Toplevel(root)
    guide_window.title("手动选择窗口")
    guide_window.geometry("400x150")
    guide_window.attributes('-topmost', True)
    guide_window.configure(bg="#F5F5F5")
    
    # 指导文本
    tk.Label(
        guide_window, 
        text="请在点击「开始选择」后，立即切换并点击微信窗口", 
        font=("Arial", 12, "bold"),
        bg="#F5F5F5",
        wraplength=380
    ).pack(pady=(15, 5))
    
    tk.Label(
        guide_window, 
        text="点击微信窗口后，将自动记住该窗口用于发送消息", 
        font=("Arial", 10),
        bg="#F5F5F5",
        wraplength=380
    ).pack(pady=5)
    
    countdown_var = tk.StringVar(value="准备开始...")
    countdown_label = tk.Label(
        guide_window,
        textvariable=countdown_var,
        font=("Arial", 12, "bold"),
        fg="#FF9800",
        bg="#F5F5F5"
    )
    countdown_label.pack(pady=5)
    
    # 存储上一个活动窗口的句柄
    last_active_window = None
    
    # 开始选择过程
    def start_selection():
        nonlocal last_active_window
        
        try:
            # 尝试获取当前活动窗口
            import ctypes
            user32 = ctypes.windll.user32
            last_active_window = user32.GetForegroundWindow()
            
            # 开始倒计时
            for i in range(5, 0, -1):
                countdown_var.set(f"请在{i}秒内点击微信窗口...")
                guide_window.update()
                time.sleep(1)
            
            # 获取用户点击后的窗口
            new_active_window = user32.GetForegroundWindow()
            
            # 如果用户没有切换窗口，提示错误
            if new_active_window == guide_window.winfo_id() or new_active_window == last_active_window:
                countdown_var.set("未检测到窗口切换，请重试")
                return
            
            # 记录选中的窗口
            global selected_wechat_window
            selected_wechat_window = new_active_window
            
            # 尝试获取窗口标题并保存
            window_title = ""
            try:
                # 获取窗口标题的长度
                title_length = user32.GetWindowTextLengthW(new_active_window) + 1
                title_buffer = ctypes.create_unicode_buffer(title_length)
                user32.GetWindowTextW(new_active_window, title_buffer, title_length)
                window_title = title_buffer.value
                
                # 保存窗口标题供下次使用
                global selected_window_title
                selected_window_title = window_title
                
                # 保存到用户设置
                if window_title and window_title not in USER_SETTINGS["saved_windows"]:
                    USER_SETTINGS["saved_windows"].insert(0, window_title)
                    # 只保留最近的5个窗口
                    USER_SETTINGS["saved_windows"] = USER_SETTINGS["saved_windows"][:5]
                    save_user_settings()
                
                log_message(f"已手动选择窗口: {window_title} (hwnd: {new_active_window})")
                show_notification(f"已选择窗口: {window_title}", "success")
            except:
                log_message(f"已手动选择窗口 (hwnd: {new_active_window})")
                show_notification("已手动选择窗口", "success")
            
            # 关闭指导窗口
            guide_window.destroy()
            
            # 重新激活主窗口
            root.focus_force()
            
        except Exception as e:
            log_message(f"手动选择窗口失败: {e}")
            countdown_var.set(f"选择失败: {e}")
    
    # 按钮区域
    button_frame = tk.Frame(guide_window, bg="#F5F5F5")
    button_frame.pack(pady=10)
    
    # 开始选择按钮
    start_button = tk.Button(
        button_frame,
        text="开始选择",
        command=start_selection,
        bg="#4CAF50",
        fg="white",
        font=("Arial", 10, "bold"),
        padx=15,
        pady=5
    )
    start_button.pack(side=tk.LEFT, padx=10)
    
    # 取消按钮
    cancel_button = tk.Button(
        button_frame,
        text="取消",
        command=guide_window.destroy,
        bg="#F44336",
        fg="white",
        font=("Arial", 10, "bold"),
        padx=15,
        pady=5
    )
    cancel_button.pack(side=tk.LEFT, padx=10)
    
    # 设置窗口关闭事件处理
    guide_window.protocol("WM_DELETE_WINDOW", guide_window.destroy)
    
    # 窗口居中显示
    guide_window.update_idletasks()
    width = guide_window.winfo_width()
    height = guide_window.winfo_height()
    x = (guide_window.winfo_screenwidth() // 2) - (width // 2)
    y = (guide_window.winfo_screenheight() // 2) - (height // 2)
    guide_window.geometry(f"{width}x{height}+{x}+{y}")
    
    return True

# 新增函数：通过窗口标题查找窗口
def find_window_by_title(title):
    """通过窗口标题查找窗口句柄"""
    if not title:
        return None
        
    # 尝试使用ctypes查找窗口
    try:
        import ctypes
        user32 = ctypes.windll.user32
        
        # 定义回调函数类型
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.POINTER(ctypes.c_int))
        
        # 存储找到的窗口句柄
        found_hwnd = [None]
        
        # 定义回调函数
        def enum_windows_callback(hwnd, lParam):
            # 获取窗口标题
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buffer, length + 1)
                window_title = buffer.value
                
                # 检查是否匹配目标标题
                if title in window_title:
                    found_hwnd[0] = hwnd
                    return False  # 停止枚举
            return True  # 继续枚举
        
        # 枚举所有窗口
        user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)
        
        return found_hwnd[0]
    except:
        # 尝试使用win32gui查找窗口
        try:
            import win32gui
            
            def callback(hwnd, param):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if title in window_title:
                        param.append(hwnd)
                return True
            
            found = []
            win32gui.EnumWindows(callback, found)
            
            return found[0] if found else None
        except:
            return None

# 恢复上次选择的窗口
def restore_saved_window():
    """尝试恢复上次保存的窗口"""
    global selected_wechat_window, selected_window_title
    
    if not USER_SETTINGS["saved_windows"]:
        return False
    
    # 遍历保存的窗口标题，尝试查找匹配的窗口
    for window_title in USER_SETTINGS["saved_windows"]:
        hwnd = find_window_by_title(window_title)
        if hwnd:
            selected_wechat_window = hwnd
            selected_window_title = window_title
            log_message(f"已恢复上次选择的窗口: {window_title}")
            return True
    
    return False

if __name__ == "__main__":
    # 全局线程变量
    monitor_thread = None
    main() 