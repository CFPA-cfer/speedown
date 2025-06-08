import os
import sys
import time
import json
import threading
import subprocess
import platform
import psutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import aria2p


class Aria2Controller:
    def __init__(self, rpc_secret="my_secret", port=6800, config_path=None):
        self.rpc_secret = rpc_secret
        self.port = port
        self.config_path = config_path
        self.aria2_process = None
        self.aria2_executable = self._find_aria2_executable()
        self.aria2_running = False

    def _find_aria2_executable(self):
        """查找系统上的Aria2可执行文件路径"""
        # 检查常见安装路径
        possible_paths = []

        if platform.system() == "Windows":
            possible_paths = [
                "C:\\Program Files\\Aria2\\aria2c.exe",
                "C:\\aria2\\aria2c.exe",
                os.path.expanduser("~\\AppData\\Local\\Programs\\aria2\\aria2c.exe")
            ]
            # 尝试在PATH中查找
            for path in os.environ["PATH"].split(";"):
                possible_paths.append(os.path.join(path, "aria2c.exe"))
        else:  # Linux/macOS
            possible_paths = [
                "/usr/bin/aria2c",
                "/usr/local/bin/aria2c",
                "/opt/homebrew/bin/aria2c"  # macOS Homebrew安装路径
            ]
            # 尝试在PATH中查找
            for path in os.environ["PATH"].split(":"):
                possible_paths.append(os.path.join(path, "aria2c"))

        # 检查路径是否存在
        for path in possible_paths:
            if os.path.exists(path):
                return path

        return None

    def start_aria2(self):
        """启动Aria2 RPC服务"""
        if not self.aria2_executable:
            return False, "Aria2 executable not found"

        if self.is_aria2_running():
            return True, "Aria2 RPC service is already running"

        # 构建启动命令
        cmd = [
            self.aria2_executable,
            "--enable-rpc",
            f"--rpc-listen-port={self.port}",
            f"--rpc-secret={self.rpc_secret}",
            "--rpc-allow-origin-all",
            "--rpc-listen-all=true",
            "--daemon=true" if platform.system() != "Windows" else ""
        ]

        # 添加配置文件（如果提供）
        if self.config_path and os.path.exists(self.config_path):
            cmd.append(f"--conf-path={self.config_path}")

        try:
            # 在Windows上以分离进程启动
            if platform.system() == "Windows":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                self.aria2_process = subprocess.Popen(
                    cmd,
                    startupinfo=startupinfo,
                    creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW
                )
            else:  # Linux/macOS
                self.aria2_process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    start_new_session=True
                )

            # 等待服务启动
            time.sleep(2)

            if self.is_aria2_running():
                self.aria2_running = True
                return True, "Aria2 RPC service started successfully"
            else:
                return False, "Failed to start Aria2 RPC service"
        except Exception as e:
            return False, f"Error starting Aria2: {str(e)}"

    def stop_aria2(self):
        """停止Aria2 RPC服务"""
        if not self.is_aria2_running():
            return True, "Aria2 RPC service is not running"

        try:
            # 如果在当前进程启动的，尝试终止它
            if self.aria2_process and self.aria2_process.poll() is None:
                self.aria2_process.terminate()
                try:
                    self.aria2_process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self.aria2_process.kill()

            # 确保所有aria2c进程都被终止
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] == 'aria2c' or proc.info['name'] == 'aria2c.exe':
                    try:
                        proc.terminate()
                    except psutil.NoSuchProcess:
                        pass

            # 等待进程退出
            time.sleep(1)

            if not self.is_aria2_running():
                self.aria2_running = False
                return True, "Aria2 RPC service stopped successfully"
            else:
                return False, "Failed to stop Aria2 RPC service"
        except Exception as e:
            return False, f"Error stopping Aria2: {str(e)}"

    def is_aria2_running(self):
        """检查Aria2服务是否在运行"""
        # 检查端口占用情况
        try:
            for conn in psutil.net_connections():
                if conn.laddr.port == self.port and conn.status == 'LISTEN':
                    return True
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            pass

        # 检查进程是否存在
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'aria2c' or proc.info['name'] == 'aria2c.exe':
                return True

        return False


class Aria2DownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aria2 高速下载器")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        # 加载图标
        try:
            self.root.iconbitmap("aria2_icon.ico")
        except:
            pass

        # 创建控制器
        self.controller = Aria2Controller()
        self.aria2_client = None

        # 创建样式
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=("Arial", 10))
        self.style.configure("TLabel", padding=5, font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))

        # 创建主框架
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建标签页
        self.tab_control = ttk.Notebook(self.main_frame)

        # 下载标签页
        self.download_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.download_tab, text="下载")

        # 设置标签页
        self.settings_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.settings_tab, text="设置")

        # 状态标签页
        self.status_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.status_tab, text="状态")

        self.tab_control.pack(fill=tk.BOTH, expand=True)

        # 初始化标签页
        self.setup_download_tab()
        self.setup_settings_tab()
        self.setup_status_tab()

        # 启动Aria2服务
        self.start_aria2_service()

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 状态更新线程
        self.status_thread = None
        self.running = True

    def setup_download_tab(self):
        """设置下载标签页"""
        # 下载链接输入区域
        url_frame = ttk.LabelFrame(self.download_tab, text="下载链接")
        url_frame.pack(fill=tk.X, padx=10, pady=5)

        self.url_text = scrolledtext.ScrolledText(url_frame, height=5, wrap=tk.WORD)
        self.url_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.url_text.insert(tk.END, "在此输入下载链接（每行一个）")

        # 下载选项区域
        options_frame = ttk.LabelFrame(self.download_tab, text="下载选项")
        options_frame.pack(fill=tk.X, padx=10, pady=5)

        # 下载路径选择
        ttk.Label(options_frame, text="保存路径:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.download_path_var = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        path_entry = ttk.Entry(options_frame, textvariable=self.download_path_var, width=50)
        path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
        ttk.Button(options_frame, text="浏览...", command=self.browse_download_path).grid(row=0, column=2, padx=5,
                                                                                          pady=5)

        # 线程数设置
        ttk.Label(options_frame, text="线程数:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.threads_var = tk.StringVar(value="8")
        threads_spin = ttk.Spinbox(options_frame, from_=1, to=64, textvariable=self.threads_var, width=5)
        threads_spin.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        # 限速设置
        ttk.Label(options_frame, text="下载限速 (KB/s):").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.speed_limit_var = tk.StringVar(value="0")
        speed_entry = ttk.Entry(options_frame, textvariable=self.speed_limit_var, width=10)
        speed_entry.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        ttk.Label(options_frame, text="(0表示不限速)").grid(row=1, column=4, padx=5, pady=5, sticky=tk.W)

        # 按钮区域
        button_frame = ttk.Frame(self.download_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.download_button = ttk.Button(button_frame, text="开始下载", command=self.start_download)
        self.download_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="暂停下载", command=self.pause_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="恢复下载", command=self.resume_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空列表", command=self.clear_download_list).pack(side=tk.LEFT, padx=5)

        # 下载任务列表
        list_frame = ttk.LabelFrame(self.download_tab, text="下载任务")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("gid", "filename", "status", "progress", "speed", "size")
        self.download_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", selectmode="browse"
        )

        # 设置列标题
        self.download_tree.heading("gid", text="任务ID")
        self.download_tree.heading("filename", text="文件名")
        self.download_tree.heading("status", text="状态")
        self.download_tree.heading("progress", text="进度")
        self.download_tree.heading("speed", text="速度")
        self.download_tree.heading("size", text="大小")

        # 设置列宽
        self.download_tree.column("gid", width=80)
        self.download_tree.column("filename", width=200)
        self.download_tree.column("status", width=80)
        self.download_tree.column("progress", width=80)
        self.download_tree.column("speed", width=80)
        self.download_tree.column("size", width=80)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.download_tree.yview)
        self.download_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.download_tree.pack(fill=tk.BOTH, expand=True)

        # 下载进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.download_tab, variable=self.progress_var, maximum=100
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

    def setup_settings_tab(self):
        """设置设置标签页"""
        # Aria2配置区域
        aria2_frame = ttk.LabelFrame(self.settings_tab, text="Aria2 设置")
        aria2_frame.pack(fill=tk.X, padx=10, pady=5)

        # Aria2路径设置
        ttk.Label(aria2_frame, text="Aria2路径:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.aria2_path_var = tk.StringVar()
        path_entry = ttk.Entry(aria2_frame, textvariable=self.aria2_path_var, width=50)
        path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
        ttk.Button(aria2_frame, text="浏览...", command=self.browse_aria2_path).grid(row=0, column=2, padx=5, pady=5)

        # RPC密钥设置
        ttk.Label(aria2_frame, text="RPC密钥:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.rpc_secret_var = tk.StringVar(value="my_secret")
        ttk.Entry(aria2_frame, textvariable=self.rpc_secret_var).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        # RPC端口设置
        ttk.Label(aria2_frame, text="RPC端口:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.rpc_port_var = tk.StringVar(value="6800")
        ttk.Entry(aria2_frame, textvariable=self.rpc_port_var, width=10).grid(row=2, column=1, padx=5, pady=5,
                                                                              sticky=tk.W)

        # 服务控制按钮
        button_frame = ttk.Frame(aria2_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        self.start_button = ttk.Button(button_frame, text="启动服务", command=self.start_aria2_service)
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="停止服务", command=self.stop_aria2_service)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # 默认下载路径
        default_frame = ttk.LabelFrame(self.settings_tab, text="默认下载路径")
        default_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(default_frame, text="默认路径:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.default_path_var = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        path_entry = ttk.Entry(default_frame, textvariable=self.default_path_var, width=50)
        path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W + tk.E)
        ttk.Button(default_frame, text="浏览...", command=self.browse_default_path).grid(row=0, column=2, padx=5,
                                                                                         pady=5)

        # 保存设置按钮
        save_frame = ttk.Frame(self.settings_tab)
        save_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(save_frame, text="保存设置", command=self.save_settings).pack(pady=5)

        # 加载保存的设置
        self.load_settings()

    def setup_status_tab(self):
        """设置状态标签页"""
        # 状态信息区域
        status_frame = ttk.LabelFrame(self.status_tab, text="服务状态")
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_text = scrolledtext.ScrolledText(status_frame, height=10, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)

        # 活动任务区域
        tasks_frame = ttk.LabelFrame(self.status_tab, text="活动下载任务")
        tasks_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("gid", "filename", "status", "progress", "speed", "size")
        self.active_tree = ttk.Treeview(
            tasks_frame, columns=columns, show="headings", selectmode="browse"
        )

        # 设置列标题
        self.active_tree.heading("gid", text="任务ID")
        self.active_tree.heading("filename", text="文件名")
        self.active_tree.heading("status", text="状态")
        self.active_tree.heading("progress", text="进度")
        self.active_tree.heading("speed", text="速度")
        self.active_tree.heading("size", text="大小")

        # 设置列宽
        self.active_tree.column("gid", width=80)
        self.active_tree.column("filename", width=200)
        self.active_tree.column("status", width=80)
        self.active_tree.column("progress", width=80)
        self.active_tree.column("speed", width=80)
        self.active_tree.column("size", width=80)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(tasks_frame, orient=tk.VERTICAL, command=self.active_tree.yview)
        self.active_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.active_tree.pack(fill=tk.BOTH, expand=True)

        # 刷新按钮
        button_frame = ttk.Frame(self.status_tab)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(button_frame, text="刷新状态", command=self.update_status).pack(pady=5)

    def browse_download_path(self):
        """浏览下载路径"""
        path = filedialog.askdirectory()
        if path:
            self.download_path_var.set(path)

    def browse_aria2_path(self):
        """浏览Aria2路径"""
        if platform.system() == "Windows":
            filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        else:
            filetypes = [("All files", "*")]

        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.aria2_path_var.set(path)

    def browse_default_path(self):
        """浏览默认下载路径"""
        path = filedialog.askdirectory()
        if path:
            self.default_path_var.set(path)

    def start_aria2_service(self):
        """启动Aria2服务"""
        # 更新控制器设置
        self.controller.aria2_executable = self.aria2_path_var.get()
        self.controller.rpc_secret = self.rpc_secret_var.get()
        self.controller.port = int(self.rpc_port_var.get())

        success, message = self.controller.start_aria2()
        self.log_status(message)

        if success:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.connect_to_aria2()

            # 启动状态更新线程
            if not self.status_thread or not self.status_thread.is_alive():
                self.status_thread = threading.Thread(target=self.update_status_thread, daemon=True)
                self.status_thread.start()

    def stop_aria2_service(self):
        """停止Aria2服务"""
        success, message = self.controller.stop_aria2()
        self.log_status(message)

        if success:
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.aria2_client = None

    def connect_to_aria2(self):
        """连接到Aria2 RPC服务"""
        try:
            self.aria2_client = aria2p.API(
                aria2p.Client(
                    host="http://localhost",
                    port=self.controller.port,
                    secret=self.controller.rpc_secret
                )
            )
            self.log_status("成功连接到Aria2 RPC服务")
            return True
        except Exception as e:
            self.log_status(f"连接Aria2 RPC服务失败: {str(e)}")
            return False

    def log_status(self, message):
        """记录状态信息"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)

    def start_download(self):
        """开始下载"""
        if not self.aria2_client:
            messagebox.showerror("错误", "Aria2服务未启动或连接失败")
            return

        urls = self.url_text.get("1.0", tk.END).strip().splitlines()
        if not urls or (len(urls) == 1 and not urls[0]):
            messagebox.showwarning("警告", "请输入下载链接")
            return

        download_path = self.download_path_var.get()
        if not os.path.exists(download_path):
            os.makedirs(download_path)

        options = {
            "dir": download_path,
            "split": self.threads_var.get(),
            "max-connection-per-server": self.threads_var.get(),
        }

        speed_limit = self.speed_limit_var.get()
        if speed_limit and int(speed_limit) > 0:
            options["max-download-limit"] = f"{speed_limit}K"

        for url in urls:
            if url.strip():
                try:
                    if url.startswith("magnet:"):
                        download = self.aria2_client.add_magnet(url, options=options)
                    elif url.endswith(".torrent"):
                        download = self.aria2_client.add_torrent(url, options=options)
                    else:
                        download = self.aria2_client.add_uris([url], options=options)

                    self.add_to_download_list(download)
                    self.log_status(f"已添加下载任务: {url}")
                except Exception as e:
                    self.log_status(f"添加下载任务失败: {url} - {str(e)}")

    def add_to_download_list(self, download):
        """添加任务到下载列表"""
        self.download_tree.insert("", tk.END, values=(
            download.gid,
            download.name,
            download.status,
            f"{download.progress:.1f}%",
            f"{download.download_speed / 1024:.1f} KB/s",
            f"{download.total_length / 1024 / 1024:.1f} MB"
        ))

    def pause_download(self):
        """暂停选中的下载任务"""
        if not self.aria2_client:
            return

        selected = self.download_tree.selection()
        if not selected:
            return

        for item in selected:
            gid = self.download_tree.item(item, "values")[0]
            try:
                download = self.aria2_client.get_download(gid)
                download.pause()
                self.log_status(f"已暂停任务: {download.name}")
            except:
                pass

    def resume_download(self):
        """恢复选中的下载任务"""
        if not self.aria2_client:
            return

        selected = self.download_tree.selection()
        if not selected:
            return

        for item in selected:
            gid = self.download_tree.item(item, "values")[0]
            try:
                download = self.aria2_client.get_download(gid)
                download.resume()
                self.log_status(f"已恢复任务: {download.name}")
            except:
                pass

    def clear_download_list(self):
        """清空下载列表"""
        self.download_tree.delete(*self.download_tree.get_children())

    def update_status(self):
        """更新状态信息"""
        if self.controller.is_aria2_running():
            self.log_status("Aria2服务运行中")
        else:
            self.log_status("Aria2服务未运行")

        # 更新活动任务列表
        if self.aria2_client:
            self.active_tree.delete(*self.active_tree.get_children())
            for download in self.aria2_client.get_downloads():
                self.active_tree.insert("", tk.END, values=(
                    download.gid,
                    download.name,
                    download.status,
                    f"{download.progress:.1f}%",
                    f"{download.download_speed / 1024:.1f} KB/s",
                    f"{download.total_length / 1024 / 1024:.1f} MB"
                ))

    def update_status_thread(self):
        """状态更新线程"""
        while self.running:
            self.root.after(100, self.update_status)
            time.sleep(2)

    def save_settings(self):
        """保存设置"""
        settings = {
            "aria2_path": self.aria2_path_var.get(),
            "rpc_secret": self.rpc_secret_var.get(),
            "rpc_port": self.rpc_port_var.get(),
            "default_path": self.default_path_var.get(),
            "threads": self.threads_var.get(),
            "speed_limit": self.speed_limit_var.get()
        }

        try:
            with open("aria2_settings.json", "w") as f:
                json.dump(settings, f)
            self.log_status("设置已保存")
        except Exception as e:
            self.log_status(f"保存设置失败: {str(e)}")

    def load_settings(self):
        """加载设置"""
        try:
            if os.path.exists("aria2_settings.json"):
                with open("aria2_settings.json", "r") as f:
                    settings = json.load(f)

                self.aria2_path_var.set(settings.get("aria2_path", ""))
                self.rpc_secret_var.set(settings.get("rpc_secret", "my_secret"))
                self.rpc_port_var.set(settings.get("rpc_port", "6800"))
                self.default_path_var.set(settings.get("default_path", os.path.expanduser("~/Downloads")))
                self.download_path_var.set(settings.get("default_path", os.path.expanduser("~/Downloads")))
                self.threads_var.set(settings.get("threads", "8"))
                self.speed_limit_var.set(settings.get("speed_limit", "0"))

                self.log_status("设置已加载")
        except Exception as e:
            self.log_status(f"加载设置失败: {str(e)}")

    def on_closing(self):
        """关闭窗口时的处理"""
        self.running = False
        if self.controller.is_aria2_running():
            self.controller.stop_aria2()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = Aria2DownloaderApp(root)
    root.mainloop()