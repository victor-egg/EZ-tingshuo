import os
import sys
import re
import sqlite3
import webbrowser
import requests
import json
import psutil
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import messagebox, font
import threading
import datetime
import time
from typing import List, Set, Optional, Dict, Any

def get_font_path(font_name):
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS  # 临时解压目录
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, 'fonts', f'{font_name}.ttf')

PROGRAM_NAME = 'ETSShell.exe'
ONLINE_DATA_URL = "https://cdn.jsdelivr.net/gh/victor-egg/EZ-tingshuo@latest/online.json"
# ONLINE_DATA_URL = "http://127.0.0.1:8000/online.json"        # DEBUG
LOG_ENCODING = 'GB18030'
DB_NAME = 'localdata/ETS.db'
MAX_RETRIES = 5
FONT_FAMILY = 'HarmonyOS Sans SC Medium'
FONT_PATH = get_font_path(f"HarmonyOS_Sans_SC_Medium")

class AppState:
    """封装应用程序运行状态"""
    def __init__(self):
        self.program_version = 20250128
        self.white_versions: List[str] = []
        self.black_versions: List[str] = []
        self.open_papers: Set[str] = set()
        self.child_windows: List[tk.Toplevel] = []
        self.examination_active = False
        self.quit_examination_flag = 0
        self.program_path: Optional[str] = None
        self.log_file_handle: Optional[int] = None

app_state = AppState()

class Application(tk.Tk):
    """主应用程序类"""
    def __init__(self):
        super().__init__()
        self.title("EZ听说")
        self.topmost_var = tk.BooleanVar(value=False)
        self._setup_fonts()
        self._setup_ui()
        self._start_background_tasks()
        self.protocol("WM_DELETE_WINDOW", self._safe_exit)

    def _setup_fonts(self):
        """配置全局字体"""
        font.nametofont("TkDefaultFont").configure(family=FONT_FAMILY)
        font.nametofont("TkTextFont").configure(family=FONT_FAMILY)
        font.nametofont("TkFixedFont").configure(family=FONT_FAMILY)
        font.nametofont("TkMenuFont").configure(family=FONT_FAMILY)
        font.nametofont("TkHeadingFont").configure(family=FONT_FAMILY)
        font.nametofont("TkCaptionFont").configure(family=FONT_FAMILY)
        font.nametofont("TkSmallCaptionFont").configure(family=FONT_FAMILY)
        font.nametofont("TkIconFont").configure(family=FONT_FAMILY)
        font.nametofont("TkTooltipFont").configure(family=FONT_FAMILY)

    def _setup_ui(self):
        """初始化用户界面组件"""
        # 状态栏
        status_frame = tk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(status_frame, text="状态", width=10, anchor="w").pack(side=tk.LEFT)
        self.status_display = tk.Text(status_frame, height=1, width=30, state=tk.DISABLED)
        self.status_display.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 置顶复选框
        tk.Checkbutton(
            status_frame,
            text="置顶",
            variable=self.topmost_var,
            command=self._toggle_topmost
        ).pack(side=tk.LEFT)

        # 日志区域
        log_frame = tk.Frame(self)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_display = tk.Text(log_frame, state=tk.DISABLED)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_display.yview)
        self.log_display.config(yscrollcommand=scrollbar.set)

        self.log_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._append_log(f"Load Font: {font.Font(family=FONT_FAMILY).actual()['family']}")

    def _toggle_topmost(self):
        """切换窗口置顶状态"""
        topmost = self.topmost_var.get()
        self.attributes('-topmost', topmost)
        for window in app_state.child_windows:
            window.attributes('-topmost', topmost)

    def _start_background_tasks(self):
        """启动后台监控线程"""
        threading.Thread(target=self._check_online_updates, daemon=True).start()

    def _check_online_updates(self):
        """检查在线更新"""
        for _ in range(MAX_RETRIES):
            try:
                self._update_status(f"更新配置文件中...(由网络决定,这可能需要一点时间)")
                response = requests.get(ONLINE_DATA_URL, timeout=10)
                response.raise_for_status()
                data = response.json()
                if not data.get("allow_run", True):
                    self._safe_exit(-1)
                app_state.white_versions = data.get("white_version_list", [])
                app_state.black_versions = data.get("black_version_list", [])
                latest_version = data.get("latest_program_version", 0)
                if latest_version > app_state.program_version:
                    self._handle_update_notification(data)
                if data.get("enforcing", False):
                    self._validate_program_version()
                self._append_log(f"已更新配置文件")
                threading.Thread(target=self._monitor_program_status, daemon=True).start()
                return
            except requests.RequestException as e:
                self._append_log(f"网络请求失败: {str(e)}")
                time.sleep(2)
        messagebox.showerror("", "无法连接到服务器")
        self._safe_exit(-1)

    def _handle_update_notification(self, data: Dict[str, Any]):
        """处理更新通知"""
        update_info = (
            f"最新版本: {data['latest_program_CVersion']}\n"
            f"更新说明:\n{data['latest_program_update_info']}\n\n"
        )  
        if messagebox.askyesno("软件更新", update_info + "点击确定下载新版本"):
            webbrowser.open(data["latest_program_download_url"])
        if data.get("must_update", False):
            messagebox.showerror("", "当前版本生命周期已结束，请下载新版本")
            self._safe_exit(0)

    def _validate_program_version(self):
        """验证程序版本有效性"""
        if not app_state.program_path:
            return False
        try:
            db_path = os.path.join(app_state.program_path, DB_NAME)
            with sqlite3.connect(db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT Value FROM Config WHERE Entry='Update_Version'"
                )
                version = cursor.fetchone()[0]
                
                valid = (
                    version in app_state.white_versions and
                    version not in app_state.black_versions
                )
                self._update_status("E听说中学 - 运行中" if valid else "版本不支持")
                return valid
        except (sqlite3.Error, TypeError) as e:
            self._append_log(f"数据库错误: {str(e)}")
            return False

    def _monitor_program_status(self):
        """监控目标程序状态"""
        while True:
            app_state.program_path = self._find_program_path()
            if app_state.program_path:
                if self._validate_program_version():
                    threading.Thread(target=self._log_monitor, daemon=True).start()
                    break
                else:
                    self._update_status("E听说中学 - 版本不支持")
            else:
                self._update_status("E听说中学 - 未启动")
            time.sleep(1)

    def _find_program_path(self) -> Optional[str]:
        """查找目标程序路径"""
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                if proc.info['name'] == PROGRAM_NAME:
                    return os.path.dirname(proc.info['exe'])
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return None

    def _log_monitor(self):
        """启动日志监控"""
        today = datetime.datetime.now().strftime('%Y-%#m-%#d')
        log_path = os.path.join(app_state.program_path, 'logs', f'shell_{today}.log')
        while True:
            try:
                if not os.path.exists(log_path):
                    raise FileNotFoundError(f"无法读取文件: {log_path}")
                with open(log_path, 'r', encoding=LOG_ENCODING, errors='ignore') as f:
                    f.seek(0, os.SEEK_END)
                    while True:
                        line = f.readline()
                        if line:
                            self._process_log_line(line)
                        time.sleep(0.01)
            except Exception as e:
                self._append_log(f"日志监控错误: {str(e)}")
                time.sleep(3)

    def _process_log_line(self, line: str):
        """处理日志条目"""
        if 'filehelper::OnGetBase64' in line:
            line = line.replace('\n', '').replace('\r', '')
            self._append_log(f"监听到 [文件操作]: {line}")
            self._handle_file_operation(line)
        elif "destroy 结束" in line:
            line = line.replace('\n', '').replace('\r', '')
            self._append_log(f"监听到 [窗口更新]: {line}")
            self._handle_exam_end()

    def _handle_file_operation(self, line: str):
        """处理文件操作日志"""
        match = re.search(r'filepath:\s*(.+?)(\s*$)', line)
        if match:
            file_path = match.group(1).strip()
            if 'template' in file_path and 'zip' not in file_path:
                if app_state.quit_examination_flag != 1:
                    self._append_log(f"更新状态: {file_path}")
                    self._init_examination(file_path)
                else:
                    app_state.quit_examination_flag = 0

    def _init_examination(self, file_path: str):
        """初始化考试流程"""
        self._update_status("考试 - 进行中")
        app_state.examination_active = True
        paper_path = os.path.dirname(os.path.dirname(file_path))
        self._analyze_paper(paper_path)

    def _handle_exam_end(self):
        """处理考试结束"""
        self._update_status("考试 - 结束")
        self._close_all_child_windows()
        app_state.quit_examination_flag = 1
        app_state.examination_active = False

    def _analyze_paper(self, path: str):
        """分析试卷内容"""
        self._append_log(f"分析试卷: {path}")
        for root, _, files in os.walk(path):
            if 'content.json' in files:
                json_path = os.path.join(root, 'content.json')
                try:
                    with open(json_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        self._create_answer_window(data)
                except (json.JSONDecodeError, IOError) as e:
                    messagebox.showerror("解析错误", f"无法解析试卷: {str(e)}")

    def _create_answer_window(self, data: Dict[str, Any]):
        """创建答案窗口"""
        paper_id = data['info']['stid']
        if paper_id in app_state.open_papers:
            return
        content = self._format_content(data)
        window = self._create_window(
            title=self._get_window_title(data['structure_type']),
            content=content,
            paper_id=paper_id
        )
        app_state.child_windows.append(window)
        app_state.open_papers.add(paper_id)

    def _get_window_title(self, struct_type: str) -> str:
        """根据试卷类型获取窗口标题"""
        return {
            'collector.read': '模范朗读',
            'collector.3q5a': '角色扮演(3问5答)',
            'collector.picture': '故事复述'
        }.get(struct_type, '未知题型')

    def _format_content(self, data: Dict[str, Any]) -> str:
        """格式化试卷内容"""
        content = re.sub(r'<.*?>', '\n', data['info']['value']).strip()
        
        if data['structure_type'] == 'collector.3q5a':
            content = f"情景:\n{content}\n\n\n" + self._format_qa_content(data)
        
        return content

    def _format_qa_content(self, data: Dict[str, Any]) -> str:
        """格式化问答内容"""
        result = []
        for idx, qa in enumerate(data['info']['question'], 1):
            result.extend([
                f"{idx}",
                f"提问(对话):  {qa['ask']}",
                f"回答(对话):  {qa['answer']}",
                "答案:"
            ])
            result.extend(f"- {std['value']}" for std in qa['std'])
            result.append(f"关键字: {qa['keywords']}\n\n")
        return '\n'.join(result)

    def _create_window(self, title: str, content: str, paper_id: str) -> tk.Toplevel:
        """创建子窗口"""
        window = tk.Toplevel(self)
        window.title(title)
        font_size = tk.IntVar(value=12)
        
        # 控制面板
        control_frame = tk.Frame(window)
        control_frame.pack(fill=tk.X)
        
        def adjust_size(delta):
            new_size = max(8, font_size.get() + delta)
            font_size.set(new_size)
            # 直接更新字体配置，避免重复创建 Font 对象
            text_box.config(font=(FONT_FAMILY, new_size))
        
        tk.Button(control_frame, text="放大", command=lambda: adjust_size(2)).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(control_frame, text="缩小", command=lambda: adjust_size(-2)).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 文本显示区域
        text_frame = tk.Frame(window)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # 初始化字体对象（推荐提前全局初始化）
        text_font = font.Font(family=FONT_FAMILY, size=font_size.get())
        
        text_box = tk.Text(
            text_frame, 
            wrap=tk.WORD, 
            font=text_font,
            state=tk.DISABLED  # 初始禁用
        )
        scroll = tk.Scrollbar(text_frame, command=text_box.yview)
        text_box.config(yscrollcommand=scroll.set)
        
        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 插入内容前临时启用写入
        text_box.config(state=tk.NORMAL)
        text_box.delete("1.0", tk.END)
        text_box.insert(tk.END, content)
        text_box.config(state=tk.DISABLED)
        
        window.protocol("WM_DELETE_WINDOW", lambda: self._close_child_window(window, paper_id))
        return window

    def _close_child_window(self, window: tk.Toplevel, paper_id: str):
        """关闭子窗口时清理资源"""
        # 移除窗口引用前先销毁组件
        for widget in window.winfo_children():
            widget.destroy()
        
        app_state.open_papers.discard(paper_id)
        if window in app_state.child_windows:
            app_state.child_windows.remove(window)
        window.destroy()

    def _close_all_child_windows(self):
        """关闭所有子窗口"""
        for window in app_state.child_windows:
            window.destroy()
        app_state.child_windows.clear()
        app_state.open_papers.clear()

    def _update_status(self, message: str):
        """更新状态栏"""
        self.status_display.config(state=tk.NORMAL)
        self.status_display.delete(1.0, tk.END)
        self.status_display.insert(tk.END, message)
        self.status_display.config(state=tk.DISABLED)

    def _append_log(self, message: str):
        """添加日志条目"""
        self.log_display.config(state=tk.NORMAL)
        self.log_display.insert(tk.END, f"[EZ听说] - {message}\n")
        self.log_display.see(tk.END)
        self.log_display.config(state=tk.DISABLED)

    def _safe_exit(self, status: int = 0):
        """安全退出程序"""
        self._close_all_child_windows()
        self.destroy()
        ctypes.windll.gdi32.RemoveFontResourceExW(FONT_PATH, 0x10, 0)
        sys.exit(status)

def main():
    """程序入口"""
    if os.name == 'nt':
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        ctypes.windll.gdi32.AddFontResourceExW.restype = wintypes.HANDLE
        ctypes.windll.gdi32.AddFontResourceExW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPVOID]
        # FR_PRIVATE (0x10) 表示仅对当前进程生效
        ctypes.windll.gdi32.AddFontResourceExW(FONT_PATH, 0x10, 0)
        app = Application()
        app.update()
        app.mainloop()
    else:
        messagebox.showerror("", "不受支持的系统")
        sys.exit(-1)

if __name__ == '__main__':
    main()
