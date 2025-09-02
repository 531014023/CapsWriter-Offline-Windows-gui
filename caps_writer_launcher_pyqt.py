import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os
import threading
import psutil
import time
import multiprocessing
import sys
from pathlib import Path
from PyQt5 import QtWidgets, QtGui, QtCore

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
        self.root.geometry("600x450")
        self.root.resizable(True, True)
        # self.root.protocol('WM_DELETE_WINDOW', self.minimize_to_tray)
        self.root.protocol('WM_DELETE_WINDOW', self.exit_application)
        
        # 初始化 Qt 应用（用于系统托盘）
        self.qt_app = QtWidgets.QApplication.instance()
        if not self.qt_app:
            self.qt_app = QtWidgets.QApplication([])
        
        # 系统托盘图标
        self.tray_icon = None
        
        # 设置 CapsWriter-Offline 路径
        self.caps_writer_dir = r".\CapsWriter-Offline-Windows-64bit"
        absolute_caps_writer_dir = os.path.abspath(self.caps_writer_dir)
        self.start_all_vbs = os.path.join(absolute_caps_writer_dir, "start_all.vbs")
        
        # 进程信息
        self.vbs_process = None
        self.is_running = False
        
        self.create_ui()
        self.create_tray_icon()
        self.check_programs_exist()
        
        # 启动时自动检查进程状态
        self.root.after(100, self.async_check_process_status)
    
    def async_check_process_status(self):
        """在后台线程中异步检查进程状态"""
        threading.Thread(target=self.check_process_status, daemon=True).start()
    def check_programs_exist(self):
        """检查 start_all.vbs 脚本是否存在"""
        if not os.path.exists(self.start_all_vbs):
            messagebox.showerror("错误", f"找不到 start_all.vbs\n请检查路径: {self.caps_writer_dir}")
            self.start_btn.config(state=tk.DISABLED)
            self.log(f"错误：找不到 {self.start_all_vbs}")
        else:
            self.log(f"找到启动脚本: {self.start_all_vbs}")
    
    def create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="CapsWriter-Offline 启动器", 
                               font=("Arial", 16, "bold"), foreground="darkblue")
        title_label.pack(pady=(0, 20))
        
        # 状态显示
        self.status_var = tk.StringVar(value="状态: 就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                font=("Arial", 11), foreground="green")
        status_label.pack(pady=(0, 15))
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        self.start_btn = ttk.Button(button_frame, text="启动 CapsWriter", 
                                   command=self.start_caps_writer, width=15)
        self.start_btn.grid(row=0, column=0, padx=5)
        
        self.stop_btn = ttk.Button(button_frame, text="停止 CapsWriter", 
                                  command=self.stop_caps_writer, state=tk.DISABLED, width=15)
        self.stop_btn.grid(row=0, column=1, padx=5)
        
        minimize_btn = ttk.Button(button_frame, text="最小化到托盘", 
                                 command=self.minimize_to_tray, width=15)
        minimize_btn.grid(row=0, column=2, padx=5)
        
        check_btn = ttk.Button(button_frame, text="检查进程状态", 
                              command=self.check_process_status, width=15)
        check_btn.grid(row=0, column=3, padx=5)
        
        # 进程状态显示
        status_frame = ttk.LabelFrame(main_frame, text="进程状态监控", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Server 状态
        server_status_frame = ttk.Frame(status_frame)
        server_status_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(server_status_frame, text="Server:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.server_status_var = tk.StringVar(value="❓ 状态未知")
        server_status_label = ttk.Label(server_status_frame, textvariable=self.server_status_var,
                                       font=("Arial", 10))
        server_status_label.pack(side=tk.LEFT, padx=5)
        
        # Client 状态
        client_status_frame = ttk.Frame(status_frame)
        client_status_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(client_status_frame, text="Client:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.client_status_var = tk.StringVar(value="❓ 状态未知")
        client_status_label = ttk.Label(client_status_frame, textvariable=self.client_status_var,
                                       font=("Arial", 10))
        client_status_label.pack(side=tk.LEFT, padx=5)
        
        # 日志显示
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, height=6, width=70, font=("Consolas", 9))
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始日志
        self.log("启动器初始化完成")
        self.log(f"VBS脚本路径: {self.start_all_vbs}")
    
    def log(self, message):
        """添加日志信息（退出时禁用）"""
        timestamp = time.strftime("%H:%M:%S")
        # 在后台线程中，使用 after 方法安全地更新UI
        self.root.after(0, lambda: self._safe_log(message, timestamp))

    def _safe_log(self, message, timestamp):
        """线程安全的日志记录"""
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
    
    def check_process_status(self):
        """检查进程状态"""
        try:
            # 检查 Server 进程
            server_running, server_pid = self.check_process_by_name('start_server.exe', self.caps_writer_dir)
            if server_running:
                self.server_status_var.set(f"✅ 运行中 (PID: {server_pid})")
            else:
                self.server_status_var.set("❌ 未运行")
            
            # 检查 Client 进程
            client_running, client_pid = self.check_process_by_name('start_client.exe', self.caps_writer_dir)
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

    def check_process_by_name(self, process_name, target_dir=None):
        """按进程名检查进程"""
        try:
            for proc in psutil.process_iter(['pid', 'name', 'exe', 'cwd']):
                try:
                    if proc.info['name'].lower() == process_name.lower():
                        if target_dir is None or (
                            proc.info.get('cwd') and 
                            os.path.abspath(proc.info['cwd']) == os.path.abspath(target_dir)
                        ):
                            return True, proc.info['pid']
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            return False, None
        except Exception as e:
            self.log(f"检查进程 {process_name} 时出错: {str(e)}")
            return False, None
    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        self.tray_icon = SystemTrayIcon(self)
        self.tray_icon.show()
    
    def start_caps_writer(self):
        """启动 CapsWriter"""
        if self.is_running:
            messagebox.showinfo("信息", "CapsWriter 已经在运行中")
            return
        
        self.log("开始启动 CapsWriter...")
        self.status_var.set("状态: 启动中...")
        
        # 在单独的线程中启动进程
        threading.Thread(target=self._start_vbs_script, daemon=True).start()
    
    def _start_vbs_script(self):
        """通过VBS脚本启动CapsWriter，启动前检查进程是否存在"""
        try:
            # 1. 检查进程是否存在，而不是直接停止
            server_running, server_pid = self.check_process_by_name('start_server.exe', self.caps_writer_dir)
            client_running, client_pid = self.check_process_by_name('start_client.exe', self.caps_writer_dir)

            # 2. 根据进程状态决定后续操作
            if server_running or client_running:
                # 只有一个进程在运行，状态不一致，建议先停止
                self.log("警告：检测到Server或Client进程在运行。")
                self.log("将尝试停止现有进程后重新启动...")
                self._stop_existing_processes()
            else:
                self.log("未检测到运行中的Server或Client进程，将直接启动。")

            # 3. 启动VBS脚本
            self.log("通过VBS脚本启动CapsWriter...")
            self.log(f"执行: {self.start_all_vbs}")

            subprocess.Popen(
                ['wscript', self.start_all_vbs],
                cwd=self.caps_writer_dir,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.log("VBS脚本已启动，等待Server和Client进程初始化...")
            
            # 等待组件初始化
            for i in range(53, 0, -1):
                self.root.after(0, lambda: self.status_var.set(f"状态: 等待组件初始化...{i}秒"))
                time.sleep(1)

            self.log("检查Server和Client进程状态...")
            self.check_process_status()

            if self.is_running:
                self.root.after(0, lambda: self.status_var.set("状态: 运行中"))
                self.root.after(0, lambda: self.start_btn.config(state=tk.DISABLED))
                self.root.after(0, lambda: self.stop_btn.config(state=tk.NORMAL))
                self.log("CapsWriter 启动完成")
            else:
                self.log("警告：VBS脚本已启动，但未检测到Server和Client进程")
                self.root.after(0, lambda: self.status_var.set("状态: 启动异常"))
                
        except Exception as e:
            error_msg = f"启动VBS脚本失败: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.root.after(0, lambda: self.status_var.set("状态: 启动失败"))
    
    def _stop_existing_processes(self):
        """停止已存在的CapsWriter相关进程，包括可能的残留进程"""
        try:
            self.log("正在停止所有CapsWriter相关进程...")
            
            # 要终止的进程列表
            target_processes = ['start_server.exe', 'start_client.exe']
            
            for proc_name in target_processes:
                try:
                    # 使用taskkill命令强制终止所有指定名称的进程
                    subprocess.run(['taskkill', '/F', '/IM', proc_name], 
                                  stdout=subprocess.DEVNULL, 
                                  stderr=subprocess.DEVNULL,
                                  timeout=10)
                    self.log(f"已尝试强制终止所有 {proc_name} 进程")
                except subprocess.TimeoutExpired:
                    self.log(f"终止 {proc_name} 进程时超时")
                except Exception as e:
                    self.log(f"终止 {proc_name} 进程时出错: {str(e)}")
            
            # 额外检查：确保进程确实已终止
            # 再次检查并报告状态
            server_running, server_pid = self.check_process_by_name('start_server.exe', self.caps_writer_dir)
            client_running, client_pid = self.check_process_by_name('start_client.exe', self.caps_writer_dir)
            
            if server_running:
                self.log(f"警告: Server进程仍在运行 (PID: {server_pid})")
            else:
                self.log("Server进程已成功终止")
                
            if client_running:
                self.log(f"警告: Client进程仍在运行 (PID: {client_pid})")
            else:
                self.log("Client进程已成功终止")
                
        except Exception as e:
            self.log(f"停止进程时发生错误: {str(e)}")
    
    def stop_caps_writer(self):
        """停止 CapsWriter"""
        
        threading.Thread(target=self._stop_processes, daemon=True).start()
    
    def _stop_processes(self):
        """实际停止进程的方法"""
        self.log("正在停止 CapsWriter...")
        self.root.after(0, lambda: self.status_var.set("状态: 停止中..."))
        try:
            self._stop_existing_processes()
            self.is_running = False
            
            self.root.after(0, lambda: self.status_var.set("状态: 已停止"))
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            
            self.log("CapsWriter 已完全停止")
            self.root.after(100, self.check_process_status)
            
        except Exception as e:
            error_msg = f"停止失败: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            
    def exit_application(self, icon=None, item=None):
        """退出程序，停止所有相关进程"""
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

def main():
    """主函数"""
    root = tk.Tk()
    app = CapsWriterLauncher(root)
    
    # 居中显示窗口
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 2
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
