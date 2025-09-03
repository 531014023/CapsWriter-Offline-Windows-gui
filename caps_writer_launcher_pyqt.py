import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os
import threading
import psutil
import time
import multiprocessing
from PyQt5 import QtWidgets, QtGui
import keyboard
import pyautogui
import win32gui
import win32process

class VoiceInputIndicator:
    def __init__(self, parent):
        self.parent = parent
        self.indicator_window = None
        self.shortcut_key = None
        self.is_listening = False
        self.current_cursor_position = None
        
        # 延迟加载配置
        self.parent.root.after(1000, self.load_shortcut_config)
    
    def load_shortcut_config(self):
        """从config.py加载快捷键配置"""
        try:
            config_path = os.path.join(self.parent.caps_writer_dir, 'config.py')
            if not os.path.exists(config_path):
                self.parent.log(f"未找到配置文件: {config_path}")
                return
                
            # 动态导入config模块
            import sys
            sys.path.insert(0, self.parent.caps_writer_dir)
            import config
            
            # 获取快捷键配置
            self.shortcut_key = config.ClientConfig.shortcut
            self.parent.log(f"加载快捷键: {self.shortcut_key}")

            # self.start_shortcut_listener()
            
        except Exception as e:
            self.parent.log(f"加载快捷键配置失败: {str(e)}")
    
    def start_shortcut_listener(self):
        """开始监听快捷键"""
        try:
            # 监听快捷键按下事件
            keyboard.on_press_key(self.shortcut_key, self._on_shortcut_pressed)
            # 监听快捷键释放事件
            keyboard.on_release_key(self.shortcut_key, self._on_shortcut_released)
            
            self.parent.log(f"开始监听快捷键: {self.shortcut_key}")
            
        except ImportError:
            self.parent.log("未安装keyboard库，无法监听快捷键")
        except Exception as e:
            self.parent.log(f"快捷键监听失败: {str(e)}")
    
    def _on_shortcut_pressed(self, event):
        """快捷键按下时的回调"""
        if not self.is_listening:
            self.is_listening = True
            # 获取当前文本光标位置
            self._get_text_cursor_position()
            # 显示指示器
            self.parent.root.after(0, self.show_voice_input_indicator)
    
    def _on_shortcut_released(self, event):
        """快捷键释放时的回调"""
        if self.is_listening:
            self.is_listening = False
            self.parent.root.after(0, self.hide_voice_input_indicator)
    
    def _get_text_cursor_position(self):
        """获取当前活动窗口的文本光标位置"""
        try:
            # 获取前台窗口
            hwnd = win32gui.GetForegroundWindow()
            
            # 获取窗口线程ID和进程ID
            thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)

            # 获取桌面窗口的线程ID（修正这里）
            desktop_hwnd = win32gui.GetDesktopWindow()
            desktop_thread_id, desktop_pid = win32process.GetWindowThreadProcessId(desktop_hwnd)
            
            # 附加到目标线程
            win32process.AttachThreadInput(desktop_thread_id, thread_id, True)
            
            try:
                # 获取光标位置
                cursor_pos = win32gui.GetCaretPos()
                
                # 转换到屏幕坐标
                point = win32gui.ClientToScreen(hwnd, cursor_pos)
                self.current_cursor_position = point
                
                self.parent.log(f"文本光标位置: {point}")
                
            finally:
                # 分离线程
                win32process.AttachThreadInput(desktop_thread_id, thread_id, False)
                
        except Exception as e:
            self.parent.log(f"获取文本光标位置失败: {str(e)}")
            # 备用方案：使用鼠标位置
            x, y = pyautogui.position()
            self.current_cursor_position = (x, y)
    
    def show_voice_input_indicator(self):
        """显示语音输入指示器"""
        try:
            if self.current_cursor_position is None:
                self.parent.log("无法获取光标位置，使用默认位置")
                return
                
            x, y = self.current_cursor_position
            
            # 创建或更新提示窗口
            if self.indicator_window is None:
                self.indicator_window = tk.Toplevel(self.parent.root)
                self.indicator_window.overrideredirect(True)  # 无边框
                self.indicator_window.attributes('-alpha', 0.8)  # 半透明
                self.indicator_window.attributes('-topmost', True)  # 置顶
                self.indicator_window.configure(bg='black')
                
                # 创建标签
                label = tk.Label(self.indicator_window, 
                               text="语音输入中...", 
                               font=("Arial", 12, "bold"),
                               bg="black",
                               fg="white",
                               padx=10,
                               pady=5)
                label.pack()
            
            # 更新位置和显示
            self.indicator_window.geometry(f"+{x+15}+{y+15}")
            self.indicator_window.deiconify()
            
        except Exception as e:
            self.parent.log(f"显示语音输入指示器失败: {str(e)}")
    
    def hide_voice_input_indicator(self):
        """隐藏语音输入指示器"""
        if self.indicator_window:
            try:
                self.indicator_window.withdraw()
            except:
                pass
    
    def cleanup(self):
        """清理资源"""
        self.hide_voice_input_indicator()
        try:
            keyboard.unhook_all()
        except:
            pass

class SystemTrayIcon(QtWidgets.QSystemTrayIcon):
    def __init__(self, parent_app, icon_path=None):
        super().__init__()
        self.parent_app = parent_app
        
        # 创建托盘图标
        if icon_path and os.path.exists(icon_path):
            self.setIcon(QtGui.QIcon(icon_path))
        else:
            # 创建默认图标
            pixmap = QtGui.QPixmap(64, 64)
            pixmap.fill(QtGui.QColor(40, 40, 40))
            painter = QtGui.QPainter(pixmap)
            painter.fillRect(16, 16, 32, 32, QtGui.QColor(0, 120, 215))
            painter.end()
            self.setIcon(QtGui.QIcon(pixmap))
        
        self.setToolTip("CapsWriter 启动器")
        
        # 创建右键菜单
        self.menu = QtWidgets.QMenu()
        
        self.show_action = QtWidgets.QAction("显示主窗口", self)
        self.show_action.triggered.connect(self.parent_app.show_window)
        self.menu.addAction(self.show_action)
        
        self.start_action = QtWidgets.QAction("启动 CapsWriter", self)
        self.start_action.triggered.connect(self.parent_app.start_caps_writer)
        self.menu.addAction(self.start_action)
        
        self.stop_action = QtWidgets.QAction("停止 CapsWriter", self)
        self.stop_action.triggered.connect(self.parent_app.stop_caps_writer)
        self.menu.addAction(self.stop_action)
        
        self.check_action = QtWidgets.QAction("检查进程状态", self)
        self.check_action.triggered.connect(self.parent_app.check_process_status)
        self.menu.addAction(self.check_action)
        
        self.menu.addSeparator()
        
        self.exit_action = QtWidgets.QAction("退出", self)
        self.exit_action.triggered.connect(self.parent_app.exit_application)
        self.menu.addAction(self.exit_action)
        
        self.setContextMenu(self.menu)
        
        # 连接双击事件
        self.activated.connect(self.on_tray_activated)

    def on_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            self.parent_app.show_window()

class CapsWriterLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("CapsWriter-Offline 启动器")
        self.root.geometry("600x600")
        self.root.resizable(True, True)
        self.root.protocol('WM_DELETE_WINDOW', self.minimize_to_tray)
        # self.root.protocol('WM_DELETE_WINDOW', self.exit_application)
        
        # 初始化 Qt 应用（用于系统托盘）
        self.qt_app = QtWidgets.QApplication.instance()
        if not self.qt_app:
            self.qt_app = QtWidgets.QApplication([])
        
        # 系统托盘图标
        self.tray_icon = None
        self.server_exe = 'start_server.exe'
        self.client_exe = 'start_client.exe'
        self.target_processes = [self.server_exe, self.client_exe]
        
        # 设置 CapsWriter-Offline 路径
        self.caps_writer_dir = r".\CapsWriter-Offline-Windows-64bit"
        absolute_caps_writer_dir = os.path.abspath(self.caps_writer_dir)
        
        # 完整路径
        self.server_exe_path = os.path.join(absolute_caps_writer_dir, self.server_exe)
        self.client_exe_path = os.path.join(absolute_caps_writer_dir, self.client_exe)
        
        # 进程信息
        self.vbs_process = None
        self.is_running = False
        self.server_process = None
        self.client_process = None
        self.server_thread = None
        self.client_thread = None
        
        self.create_ui()
        self.create_tray_icon()
        self.check_programs_exist()
        
        # 在初始化完成后居中窗口
        self.center_window()
        
        # 启动时自动检查进程状态
        self.root.after(100, self.async_check_process_status)

        # 新增：语音输入状态提示相关变量
        self.voice_input_window = None
        self.is_listening = False
        
        # 添加语音输入指示器
        self.voice_indicator = VoiceInputIndicator(self)
    
    def async_check_process_status(self):
        """在后台线程中异步检查进程状态"""
        threading.Thread(target=self.check_process_status, daemon=True).start()

    def check_programs_exist(self):
        """检查 start_server.exe,start_client.exe 脚本是否存在"""
        if not os.path.exists(self.server_exe_path):
            messagebox.showerror("错误", f"找不到 {self.server_exe_path}\n请检查路径: {self.caps_writer_dir}")
            self.start_btn.config(state=tk.DISABLED)
            self.log(f"错误：找不到 {self.server_exe_path}")
        else:
            self.log(f"找到启动脚本: {self.server_exe_path}")
        
        if not os.path.exists(self.client_exe_path):
            messagebox.showerror("错误", f"找不到 {self.client_exe_path}\n请检查路径: {self.caps_writer_dir}")
            self.start_btn.config(state=tk.DISABLED)
            self.log(f"错误：找不到 {self.client_exe_path}")
        else:
            self.log(f"找到启动脚本: {self.client_exe_path}")
    
    def create_ui(self):
        """创建用户界面"""
        # 主框架 - 使用 grid 布局
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 配置主框架的网格权重
        main_frame.grid_rowconfigure(4, weight=1)  # 进程状态监控区域
        main_frame.grid_rowconfigure(6, weight=1)  # 日志区域
        main_frame.grid_columnconfigure(0, weight=1)
        
        # 标题 (row 0) - 使用Frame实现真正居中
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, pady=(0, 10), sticky="nsew")
        title_frame.grid_columnconfigure(0, weight=1)
        
        title_label = ttk.Label(title_frame, text="CapsWriter-Offline 启动器", 
                            font=("Arial", 16, "bold"), foreground="darkblue")
        title_label.grid(row=0, column=0, sticky="")  # 空sticky使其在单元格中居中[7](@ref)
        
        # 状态显示 (row 1) - 使用Frame实现真正居中
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=1, column=0, pady=(0, 10), sticky="nsew")
        status_frame.grid_columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="状态: 就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                font=("Arial", 11), foreground="green")
        status_label.grid(row=0, column=0, sticky="")  # 空sticky使其在单元格中居中[7](@ref)
        
        # 按钮框架 (row 2)
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        self.start_btn = ttk.Button(button_frame, text="启动 CapsWriter", 
                                command=self.start_caps_writer)
        self.start_btn.grid(row=0, column=0, padx=5, sticky="ew")
        
        self.stop_btn = ttk.Button(button_frame, text="停止 CapsWriter", 
                                command=self.stop_caps_writer, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5, sticky="ew")
        
        minimize_btn = ttk.Button(button_frame, text="最小化到托盘", 
                                command=self.minimize_to_tray)
        minimize_btn.grid(row=0, column=2, padx=5, sticky="ew")
        
        check_btn = ttk.Button(button_frame, text="检查进程状态", 
                            command=self.check_process_status)
        check_btn.grid(row=0, column=3, padx=5, sticky="ew")
        
        # 进程状态监控标签 (row 3)
        process_status_frame = ttk.Frame(main_frame)
        process_status_frame.grid(row=3, column=0, pady=(10, 5), sticky="nsew")
        process_status_frame.grid_columnconfigure(0, weight=1)
        
        process_status_label = ttk.Label(process_status_frame, text="进程状态监控", 
                            font=("Arial", 12, "bold"))
        process_status_label.grid(row=3, column=0, sticky="")  # 空sticky使其在单元格中居中[7](@ref)
        
        # 进程状态详情框架 (row 4)
        status_detail_frame = ttk.Frame(main_frame)
        status_detail_frame.grid(row=4, column=0, pady=5, sticky="nsew")
        status_detail_frame.grid_rowconfigure(2, weight=1)  # Server输出区域
        status_detail_frame.grid_rowconfigure(4, weight=1)  # Client输出区域
        status_detail_frame.grid_columnconfigure(0, weight=1)
        
        # Server 状态 (row 0)
        server_status_frame = ttk.Frame(status_detail_frame)
        server_status_frame.grid(row=0, column=0, pady=2, sticky="w")
        ttk.Label(server_status_frame, text="Server:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.server_status_var = tk.StringVar(value="❓ 状态未知")
        ttk.Label(server_status_frame, textvariable=self.server_status_var, font=("Arial", 10)).pack(side=tk.LEFT)
        # Client 状态 (row 1)
        client_status_frame = ttk.Frame(status_detail_frame)
        client_status_frame.grid(row=1, column=0, pady=2, sticky="w")
        ttk.Label(client_status_frame, text="Client:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.client_status_var = tk.StringVar(value="❓ 状态未知")
        ttk.Label(client_status_frame, textvariable=self.client_status_var, font=("Arial", 10)).pack(side=tk.LEFT)
                
        # Server 输出区域 (row 2)
        server_output_frame = ttk.LabelFrame(status_detail_frame, text="Server 输出")
        server_output_frame.grid(row=2, column=0, pady=5, sticky="nsew")
        
        server_text_frame = ttk.Frame(server_output_frame)
        server_text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.server_output_text = tk.Text(server_text_frame, height=4, font=("Consolas", 9))
        server_scrollbar = ttk.Scrollbar(server_text_frame, orient=tk.VERTICAL, command=self.server_output_text.yview)
        self.server_output_text.configure(yscrollcommand=server_scrollbar.set)
        
        self.server_output_text.grid(row=0, column=0, sticky="nsew")
        server_scrollbar.grid(row=0, column=1, sticky="ns")
        server_text_frame.grid_rowconfigure(0, weight=1)
        server_text_frame.grid_columnconfigure(0, weight=1)
        
        # Client 输出区域 (row 3)
        client_output_frame = ttk.LabelFrame(status_detail_frame, text="Client 输出")
        client_output_frame.grid(row=3, column=0, pady=5, sticky="nsew")
        
        client_text_frame = ttk.Frame(client_output_frame)
        client_text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.client_output_text = tk.Text(client_text_frame, height=4, font=("Consolas", 9))
        client_scrollbar = ttk.Scrollbar(client_text_frame, orient=tk.VERTICAL, command=self.client_output_text.yview)
        self.client_output_text.configure(yscrollcommand=client_scrollbar.set)
        
        self.client_output_text.grid(row=0, column=0, sticky="nsew")
        client_scrollbar.grid(row=0, column=1, sticky="ns")
        client_text_frame.grid_rowconfigure(0, weight=1)
        client_text_frame.grid_columnconfigure(0, weight=1)
 
        # 日志区域 (row 5)
        log_frame = ttk.LabelFrame(main_frame, text="操作日志")
        log_frame.grid(row=5, column=0, pady=5, sticky="nsew")
        
        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_text_frame, height=3, font=("Consolas", 9))
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        log_text_frame.grid_rowconfigure(0, weight=1)
        log_text_frame.grid_columnconfigure(0, weight=1)
        
        # 初始日志
        self.log("启动器初始化完成")
        self.log(f"Server路径: {self.server_exe_path}")
        self.log(f"Client路径: {self.client_exe_path}")

    
    def log(self, message):
        """添加日志信息（退出时禁用）"""
        timestamp = time.strftime("%H:%M:%S")
        # 在后台线程中，使用 after 方法安全地更新UI
        self.root.after(0, lambda: self._safe_log(message, timestamp))

    def _safe_log(self, message, timestamp):
        """线程安全的日志记录"""
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        
    def append_server_output(self, text):
        """添加Server输出文本"""
        self.root.after(0, lambda: self._append_text(self.server_output_text, text))
        
    def append_client_output(self, text):
        """添加Client输出文本"""
        self.root.after(0, lambda: self._append_text(self.client_output_text, text))
        
    def _append_text(self, text_widget, text):
        """在文本控件中添加文本并滚动到底部"""
        text_widget.insert(tk.END, text)
        text_widget.see(tk.END)
    
    def check_process_status(self):
        """检查进程状态"""
        try:
            process_status = self.check_processes_by_names(self.target_processes, self.caps_writer_dir)
            
            server_running, server_pid = process_status[self.server_exe]
            client_running, client_pid = process_status[self.client_exe]
            # 检查 Server 进程
            if server_running:
                self.server_status_var.set(f"✅ 运行中 (PID: {server_pid})")
            else:
                self.server_status_var.set("❌ 未运行")
            
            # 检查 Client 进程
            if client_running:
                self.client_status_var.set(f"✅ 运行中 (PID: {client_pid})")
            else:
                self.client_status_var.set("❌ 未运行")
 
            # 更新运行状态
            if server_running and client_running:
                self.is_running = True
                self.status_var.set("状态: 运行中")
                self.start_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.NORMAL)
            else:
                self.is_running = False
                self.status_var.set("状态: 已停止")
                self.start_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)
                
        except Exception as e:
            self.log(f"检查进程状态时出错: {str(e)}")

    def check_processes_by_names(self, target_processes, target_dir=None):
        """
        一次性检查多个进程的状态
        :param target_processes: 要检查的进程名称列表，如 ['start_server.exe', 'start_client.exe']
        :param target_dir: 目标工作目录（可选）
        :return: 字典，key为进程名，value为 (是否运行, PID) 的元组
        """
        result = {proc_name: (False, None) for proc_name in target_processes}
        
        try:
            # 一次性获取所有进程信息[1](@ref)[2](@ref)[3](@ref)
            processes = list(psutil.process_iter(['pid', 'name', 'cwd']))
            
            for proc in processes:
                try:
                    proc_info = proc.info
                    proc_name = proc_info.get('name', '').lower()
                    
                    # 检查是否是目标进程
                    if proc_name in [p.lower() for p in target_processes]:
                        # 检查工作目录是否匹配（如果指定了target_dir）
                        if target_dir is None or (
                            proc_info.get('cwd') and 
                            os.path.abspath(proc_info['cwd']) == os.path.abspath(target_dir)
                        ):
                            # 找到原始大小写的进程名
                            original_name = next((p for p in target_processes if p.lower() == proc_name), proc_name)
                            result[original_name] = (True, proc_info.get('pid'))
                            
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
            return result
            
        except Exception as e:
            self.log(f"检查进程时出错: {str(e)}")
            return result

    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        self.tray_icon = SystemTrayIcon(self)
        self.tray_icon.show()
    
    def start_caps_writer(self):
        """启动 CapsWriter"""
        if self.is_running:
            messagebox.showinfo("信息", "CapsWriter 已经在运行中")
            return
        
        # 检查可执行文件是否存在
        if not os.path.exists(self.server_exe_path) or not os.path.exists(self.client_exe_path):
            messagebox.showerror("错误", "找不到Server或Client可执行文件，请检查程序路径")
            return
        
        self.log("开始启动 CapsWriter...")
        self.status_var.set("状态: 启动中...")
        
        # 在单独的线程中启动进程
        threading.Thread(target=self._start_processes, daemon=True).start()
    
    def _start_processes(self):
        """直接启动Server和Client进程并捕获输出"""
        try:
            # 1. 检查进程是否存在，而不是直接停止
            process_status = self.check_processes_by_names(self.target_processes, self.caps_writer_dir)
            
            server_running, server_pid = process_status[self.server_exe]
            client_running, client_pid = process_status[self.client_exe]

            # 2. 根据进程状态决定后续操作
            if server_running or client_running:
                # 只有一个进程在运行，状态不一致，建议先停止
                self.log("警告：检测到Server或Client进程在运行。")
                self.log("将尝试停止现有进程后重新启动...")
                self._stop_existing_processes()
            else:
                self.log("未检测到运行中的Server或Client进程，将直接启动。")

            # 3. 直接启动Server和Client进程并捕获输出
            self.log("直接启动Server和Client进程...")
            
            # 设置环境变量以支持UTF-8编码，解决子进程编码问题
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            env['PYTHONUTF8'] = '1'
            env['PYTHONLEGACYWINDOWSCONSOLE'] = '1'  # 兼容Windows控制台
            env['LC_ALL'] = 'C.UTF-8'
            env['LANG'] = 'C.UTF-8'
            
            # 启动Server进程（修改编码处理）
            self.server_process = subprocess.Popen(
                [self.server_exe_path],
                cwd=self.caps_writer_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=False,
                creationflags=subprocess.CREATE_NO_WINDOW,
                env=env
            )
            
            self.log(f"Server进程已启动 (PID: {self.server_process.pid})")
            
            time.sleep(1)
            # 启动Client进程（修改编码处理）
            self.client_process = subprocess.Popen(
                [self.client_exe_path],
                cwd=self.caps_writer_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=False,
                creationflags=subprocess.CREATE_NO_WINDOW,
                env=env
            )
            
            self.log(f"Client进程已启动 (PID: {self.client_process.pid})")
            
            # 创建线程来读取输出
            self.server_thread = threading.Thread(target=self._read_server_output, daemon=True)
            self.server_thread.start()
            
            self.client_thread = threading.Thread(target=self._read_client_output, daemon=True)
            self.client_thread.start()
            
            self.check_process_status()
            # 启动监控线程，监听启动完成信号
            self.monitor_thread = threading.Thread(target=self._monitor_startup_completion, daemon=True)
            self.monitor_thread.start()
            
        except Exception as e:
            error_msg = f"启动进程失败: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.root.after(0, lambda: self.status_var.set("状态: 启动失败"))

    def _monitor_startup_completion(self):
        """监控启动完成状态"""
        server_started = False
        client_started = False
        max_wait_time = 120  # 最大等待时间120秒
        start_time = time.time()
        
        self.root.after(0, lambda: self.status_var.set("状态: 等待服务启动..."))
        
        while time.time() - start_time < max_wait_time:
            # 检查Server是否启动完成
            if not server_started and hasattr(self, 'server_startup_signal') and self.server_startup_signal:
                server_started = True
                self.log("Server服务已启动")
                self.root.after(0, lambda: self.status_var.set("状态: Server已启动，等待Client..."))
            
            # 检查Client是否启动完成
            if not client_started and hasattr(self, 'client_startup_signal') and self.client_startup_signal:
                client_started = True
                self.log("Client已连接成功")
                self.root.after(0, lambda: self.status_var.set("状态: Client已连接，等待Server..."))
            
            # 如果两者都启动完成
            if server_started and client_started:
                self.check_process_status()
                self.log("CapsWriter 启动完成")
                # 开始监听语音快捷键
                self.voice_indicator.start_shortcut_listener()
                return
            
            # 检查进程是否还在运行
            if (self.server_process and self.server_process.poll() is not None) or \
            (self.client_process and self.client_process.poll() is not None):
                self.log("警告：进程异常退出")
                break
            
            time.sleep(0.5)  # 每0.5秒检查一次
        
        # 超时或异常处理
        if not (server_started and client_started):
            self.log("警告：启动超时或进程异常")
            self.root.after(0, lambda: self.status_var.set("状态: 启动异常"))
            # 强制检查一次状态
            self.check_process_status()
    def _read_server_output(self):
        """读取Server进程的输出"""
        # 初始化启动信号
        self.server_startup_signal = False
        
        try:
            if self.server_process and self.server_process.stdout:
                while True:
                    line = self.server_process.stdout.readline()
                    if not line:
                        break
                    
                    # 改进的编码处理
                    decoded_line = self._safe_decode(line)
                    
                    if decoded_line:
                        self.append_server_output(decoded_line)
                        
                        # 检查是否包含启动完成的关键词
                        if any(keyword in decoded_line.lower() for keyword in 
                            ['开始服务', 'start service', 'service started', 'listening', '启动成功']):
                            self.server_startup_signal = True
                            self.log("检测到Server启动完成信号")
                
                # 检查进程是否正常退出
                return_code = self.server_process.wait()
                if return_code != 0:
                    self.log(f"Server进程异常退出，返回码: {return_code}")
        except Exception as e:
            self.log(f"读取Server输出时出错: {str(e)}")

    def _read_client_output(self):
        """读取Client进程的输出"""
        # 初始化启动信号
        self.client_startup_signal = False
        
        try:
            if self.client_process and self.client_process.stdout:
                while True:
                    line = self.client_process.stdout.readline()
                    if not line:
                        break
                    
                    # 改进的编码处理
                    decoded_line = self._safe_decode(line)
                    
                    if decoded_line:
                        self.append_client_output(decoded_line)
                        
                        # 检查是否包含连接成功的关键词
                        if any(keyword in decoded_line.lower() for keyword in 
                            ['连接成功', 'connected', 'connection established', 'connect success', '启动完成']):
                            self.client_startup_signal = True
                            self.log("检测到Client连接成功信号")
                
                # 检查进程是否正常退出
                return_code = self.client_process.wait()
                if return_code != 0:
                    self.log(f"Client进程异常退出，返回码: {return_code}")
        except Exception as e:
            self.log(f"读取Client输出时出错: {str(e)}")
    def _safe_decode(self, line):
        """安全地解码字节数据"""
        if isinstance(line, bytes):
            # 尝试多种编码，优先使用UTF-8
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'ascii']
            for encoding in encodings:
                try:
                    return line.decode(encoding)
                except UnicodeDecodeError:
                    continue
            
            # 如果所有编码都失败，使用错误处理方式
            try:
                return line.decode('utf-8', errors='replace')
            except:
                return f"[无法解码的数据: {line.hex()}]\n"
        else:
            return str(line)

    def _stop_existing_processes(self):
        """停止已存在的CapsWriter相关进程，包括可能的残留进程"""
        try:
            self.log("正在停止所有CapsWriter相关进程...")
            
            # 要终止的进程列表
            all_processes = list(psutil.process_iter(['pid', 'name', 'cwd'])) # 一次性获取所有进程信息

            for proc in all_processes:
                try:
                    proc_info = proc.info
                    pid = proc_info.get('pid', '')
                    proc_name = proc_info.get('name', '').lower()
                    proc_cwd = proc_info.get('cwd', '')

                    # 检查是否是目标进程且在目标目录下
                    if (proc_name in self.target_processes and proc_cwd and os.path.abspath(proc_cwd) == os.path.abspath(self.caps_writer_dir)):
                        self.log(f"尝试停止进程 {proc_name} (PID: {pid})")
                        # 尝试优雅终止
                        process = psutil.Process(pid)
                        process.terminate()  # 发送终止信号
                        # 等待进程退出
                        try:
                            process.wait(timeout=3)
                            self.log(f"进程 {proc_name} (PID: {pid}) 已优雅退出")
                        except psutil.TimeoutExpired:
                            # 优雅终止失败，强制终止
                            process.kill()
                            self.log(f"进程 {proc_name} (PID: {pid}) 已强制终止")

                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    self.log(f"终止进程时出错: {str(e)}")
            
            # 额外检查：确保进程确实已终止
            # 再次检查并报告状态
            process_status = self.check_processes_by_names(self.target_processes, self.caps_writer_dir)
            
            server_running, server_pid = process_status[self.server_exe]
            client_running, client_pid = process_status[self.client_exe]
            server_result = False
            client_result = False
            if server_running:
                self.log(f"警告: Server进程仍在运行 (PID: {server_pid})")
            else:
                self.log("Server进程已成功终止")
                server_result = True
                
            if client_running:
                self.log(f"警告: Client进程仍在运行 (PID: {client_pid})")
            else:
                self.log("Client进程已成功终止")
                client_result = True
            if server_result and client_result:
                self.log("所有CapsWriter进程已停止")
                return True
            else:
                self.log("未成功停止所有CapsWriter进程")
                return False
        except Exception as e:
            self.log(f"停止进程时发生错误: {str(e)}")
            return False
    
    def stop_caps_writer(self):
        """停止 CapsWriter"""
        
        threading.Thread(target=self._stop_processes, daemon=True).start()
    
    def _stop_processes(self):
        """实际停止进程的方法"""
        self.log("正在停止 CapsWriter...")
        self.root.after(0, lambda: self.status_var.set("状态: 停止中..."))
        try:
            result = self._stop_existing_processes()
            # 清理进程引用
            self.server_process = None
            self.client_process = None
            self.server_thread = None
            self.client_thread = None
            
            if result:
                self.is_running = False
                
                self.root.after(0, lambda: self.status_var.set("状态: 已停止"))
                self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
                
                self.log("CapsWriter 已完全停止")
                self.root.after(100, self.check_process_status)
            else:
                self.root.after(0, lambda: messagebox.showerror("❌ 停止进程失败，请检查"))
        except Exception as e:
            error_msg = f"停止失败: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            
    def exit_application(self, icon=None, item=None):
        """退出程序，停止所有相关进程"""
        # 清理语音输入指示器
        if hasattr(self, 'voice_indicator'):
            self.voice_indicator.cleanup()
            pass

        self.log("正在退出程序...")
        # 创建并启动一个线程来停止进程
        stop_thread = threading.Thread(target=self._stop_processes_and_quit, daemon=True)
        stop_thread.start()
    
    def _stop_processes_and_quit(self):
        """在后台线程中停止进程并退出应用"""
        try:
            self._stop_processes()  # 同步停止进程
        except Exception as e:
            self.log(f"停止进程时发生错误: {str(e)}")
        finally:
            # 无论停止是否成功，都继续执行退出操作
            # 使用 after 确保在主线程中执行GUI操作
            self.root.after(0, self._final_quit)

    def _final_quit(self):
        """执行最终的退出操作（在主线程中调用）"""
        # 停止托盘图标
        if self.tray_icon:
            self.tray_icon.hide()
        # 销毁主窗口
        self.root.destroy()
        # 退出Qt应用
        if self.qt_app:
            self.qt_app.quit()
    def minimize_to_tray(self):
        """最小化到系统托盘"""
        self.root.withdraw()
        self.log("程序已最小化到系统托盘")
    
    def show_window(self, icon=None, item=None):
        """从托盘显示主窗口"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.root.after(100, self._deferred_window_setup)

    def _deferred_window_setup(self):
        """显示主窗口"""
        # 更新状态（分步执行，避免阻塞）
        self.root.after(100, self.check_process_status)
        self.log("主窗口已显示")
        
        # 强制重绘界面
        self.root.update_idletasks()
        self.root.update()

    def center_window(self):
        """使窗口在屏幕上居中"""
        self.root.update_idletasks()  # 确保获取到准确的窗口尺寸
        
        # 获取屏幕尺寸和窗口尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # 设置窗口位置
        self.root.geometry(f"+{x}+{y}")

def main():
    """主函数"""
    root = tk.Tk()
    app = CapsWriterLauncher(root)
    
    root.mainloop()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()