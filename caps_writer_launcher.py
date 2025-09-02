import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os
import threading
from PIL import Image, ImageDraw
import pystray
import psutil
import time
import multiprocessing
import sys
from pathlib import Path 

class CapsWriterLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("CapsWriter-Offline 启动器")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        self.root.protocol('WM_DELETE_WINDOW', self.exit_application)
        self.tray_icon = None  # 初始化托盘图标变量
        self.tray_thread = None  # 初始化托盘线程变量
        self.tray_running = False  # 添加托盘运行状态标志
        
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
        """添加日志信息"""
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
        image = Image.new('RGB', (64, 64), (40, 40, 40))
        dc = ImageDraw.Draw(image)
        dc.rectangle((16, 16, 48, 48), fill=(0, 120, 215))
        
        menu = pystray.Menu(
            pystray.MenuItem('显示主窗口', self.show_window),
            pystray.MenuItem('启动 CapsWriter', self.start_caps_writer),
            pystray.MenuItem('停止 CapsWriter', self.stop_caps_writer),
            pystray.MenuItem('检查进程状态', self.check_process_status),
            pystray.MenuItem('退出', self.exit_application)
        )
        
        self.tray_icon = pystray.Icon("caps_writer_launcher", image, "CapsWriter 启动器", menu)
        
        # 尝试设置默认项目来响应点击（虽然不是真正的双击，但更稳定）
        # 有些版本的 pystray 支持设置默认菜单项来响应左键单击
        # 这可以作为双击的一种替代方案
        try:
            # 检查 pystray 版本是否支持 default 参数
            default_item = pystray.MenuItem('显示主窗口', self.show_window, default=True)
            # 如果支持，重新创建包含默认项的菜单
            menu = pystray.Menu(
                default_item,
                pystray.MenuItem('启动 CapsWriter', self.start_caps_writer),
                pystray.MenuItem('停止 CapsWriter', self.stop_caps_writer),
                pystray.MenuItem('检查进程状态', self.check_process_status),
                pystray.MenuItem('退出', self.exit_application)
            )
            self.tray_icon.menu = menu
        except (TypeError, AttributeError):
            # 如果不支持 default 参数，回退到原始菜单
            self.log("当前 pystray 版本不支持默认菜单项，将使用标准菜单")
            pass
    
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
            
            # ... (后续的等待和状态检查逻辑保持不变，例如等待25秒)
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
                                  creationflags=subprocess.CREATE_NO_WINDOW,
                                  timeout=10)
                    self.log(f"已尝试强制终止所有 {proc_name} 进程")
                except subprocess.TimeoutExpired:
                    self.log(f"终止 {proc_name} 进程时超时")
                except Exception as e:
                    self.log(f"终止 {proc_name} 进程时出错: {str(e)}")
            
            
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
            self.root.after(1000, self.check_process_status)
            
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
            self.tray_icon.stop()
        # 销毁主窗口
        self.root.destroy()
        # 退出应用
        os._exit(0)
    def minimize_to_tray(self):
        """最小化到系统托盘"""
        self.root.withdraw()
        
         # 如果托盘图标不存在或已停止，重新创建
        if self.tray_icon is None or not self.tray_running:
            self.create_tray_icon()  # 重新创建托盘图标
        
        self.tray_thread_start()
        
        self.log("程序已最小化到系统托盘")
        
    def tray_thread_start(self):
        # 确保只启动一次托盘图标线程
        if self.tray_thread is None or not self.tray_thread.is_alive():
            self.tray_thread = threading.Thread(target=self._run_tray_icon, daemon=True)
            self.tray_thread.start()
        
    def _run_tray_icon(self):
        """运行托盘图标（在单独线程中）"""
        try:
            if self.tray_icon:
                self.tray_running = True
                self.tray_icon.run()
        except Exception as e:
            self.log(f"托盘图标运行错误: {str(e)}")
        finally:
            self.tray_running = False
    
    def show_window(self, icon=None, item=None):
        """从托盘显示主窗口"""
        # 先立即显示窗口框架，避免白屏
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        # 不要停止托盘图标，只是隐藏主窗口
        # 托盘图标会在窗口显示后继续运行
        # 延迟执行耗时操作，让窗口先显示出来
        self.root.after(100, self._deferred_window_setup)
    
    def _deferred_window_setup(self):
        """显示主窗口"""
        # 更新状态（分步执行，避免阻塞）
        self.root.after(100, self.check_process_status)
        self.log("主窗口已显示")
        
        # 停止托盘图标
        if self.tray_icon:
            self.tray_icon.stop()
        
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
