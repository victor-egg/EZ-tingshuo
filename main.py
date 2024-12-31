import os
import re
import psutil
import json
import tkinter as tk
from tkinter import messagebox
import threading
import datetime
import time

open_papers = []
open_child_windows = []
examination_status = False
m_quit_examination = 0

def toggle_topmost():
    global open_child_windows
    root.attributes('-topmost', topmost_var.get())
    for window in open_child_windows:
        window.attributes('-topmost', topmost_var.get())

def find_program_directory(program_name):
    for proc in psutil.process_iter(['name', 'exe']):
        try:
            if proc.info['name'] == program_name:
                return os.path.dirname(proc.info['exe'])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return None

def check_program_status(program_name):
    global open_child_windows
    global program_directory
    while True:
        program_directory = find_program_directory(program_name)
        if program_directory:
            update_status("E听说中学 - 运行中")
            start_log_watcher()
            break
        else:
            update_status("E听说中学 - 未启动")
        time.sleep(1)

def log_watcher():
    global examination_status
    global m_quit_examination
    today_date = datetime.datetime.now().strftime('%Y-%m-%d')
    log_file = os.path.join(program_directory, 'logs', f'shell_{today_date}.log')
    while True:
        try:
            with open(log_file, 'r', encoding='GB18030') as f:
                f.seek(0, os.SEEK_END)
                while True:
                    line = f.readline()
                    if not line:
                        time.sleep(0.1)
                        continue
                    update_log(line)

                    if 'filehelper::OnGetBase64' in line:
                        match = re.search(r'filepath:\s*(.+?)(\s*$)', line)
                        if match:
                            file_path = match.group(1).strip()
                            if 'template' in file_path and 'zip' not in file_path:
                                if m_quit_examination != 1:
                                    update_status("考试 - 进行中")
                                    examination_status = True
                                    examination_paper_path = os.path.dirname(os.path.dirname(file_path))
                                    init_paper(examination_paper_path)
                                else:
                                    m_quit_examination = 0

                    if "destroy 结束" in line:
                        update_status("考试 - 结束")
                        close_all_child_windows()
                        m_quit_examination = 1
                        examination_status = False
        except UnicodeDecodeError as e:
            update_log("[EZ听说] 程序运行错误: {e}")
        except Exception as e:
            messagebox.showerror("遇到了意料之外的情况~", f"解析日志遇到问题: {e}")

def close_all_child_windows():
    global open_papers
    global open_child_windows
    for window in open_child_windows:
        if window.winfo_exists():
            window.destroy() 
    open_papers.clear()

def init_paper(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if file == 'content.json':
                json_path = os.path.join(root, file)
                analyze_paper(json_path)

def format_string(string):
    string = re.sub(r'<.*?>', '\n', string)  # 去除HTML标签
    string = string.strip()  # 去除首尾空格
    return string

def analyze_paper(json_file_path):
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if data['structure_type'] == 'collector.read':
                paper_id = data['info']['stid']
                content = format_string(data['info']['value'])
                create_windows_paper_answer("模范朗读", content, paper_id)
            if data['structure_type'] == 'collector.3q5a':
                paper_id = data['info']['stid']
                background_talk = format_string(data['info']['value'])
                content = f"情景:\n{background_talk}\n\n\n"

                part = ""
                for i in range(len(data['info']['question'])):
                    tmp = data['info']['question'][i]
                    # 提取信息
                    question = tmp['ask']
                    answer = tmp['answer']
                    keywords = tmp['keywords']
                    std_responses = tmp['std']
                    # 拼接部分内容
                    part += f"{i+1}\n"
                    part += f"提问(对话):  {question}\n"
                    part += f"回答(对话):  {answer}\n"
                    part += "答案:\n"
                    for std in std_responses:
                        part += f"- {std['value']}\n"
                    part += f"关键字: {keywords}\n"
                    part += "\n\n"

                content += part
                create_windows_paper_answer("角色扮演(3问5答)", content, paper_id)
            if data['structure_type'] == 'collector.picture':
                paper_id = data['info']['stid']
                content = format_string(data['info']['value'])
                create_windows_paper_answer("故事复述", content, paper_id)
    except json.JSONDecodeError:
        messagebox.showerror("遇到了意料之外的情况~", f"无法解析试题: {json_file_path}")

def update_log(message):
    text_box_log.config(state=tk.NORMAL)
    text_box_log.insert(tk.END, message)
    text_box_log.see(tk.END)
    text_box_log.config(state=tk.DISABLED)

def update_status(message):
    status_box.config(state=tk.NORMAL)
    status_box.delete(1.0, tk.END)
    status_box.insert(tk.END, message)
    status_box.config(state=tk.DISABLED)

def start_log_watcher():
    thread = threading.Thread(target=log_watcher)
    thread.daemon = True
    thread.start()

def create_windows_paper_answer(title, content, paper_id):
    global open_papers
    global open_child_windows
    if paper_id in open_papers:
        return
    else:
        open_papers.append(paper_id)

    new_window = tk.Toplevel(root)
    new_window.title(title)
    open_child_windows.append(new_window)

    # 创建一个框架用于按钮和文本框
    button_frame = tk.Frame(new_window)
    button_frame.pack(fill=tk.X)

    # 创建字体大小调整函数
    def increase_font_size():
        current_size = text_box.cget("font").split()[1]
        new_size = int(current_size) + 2
        text_box.config(font=("Arial", new_size))

    def decrease_font_size():
        current_size = text_box.cget("font").split()[1]
        new_size = max(1, int(current_size) - 2)  # 确保字体大小不小于1
        text_box.config(font=("Arial", new_size))

    # 创建按钮
    increase_button = tk.Button(button_frame, text="放大", command=increase_font_size)
    increase_button.pack(side=tk.LEFT, padx=5, pady=5)

    decrease_button = tk.Button(button_frame, text="缩小", command=decrease_font_size)
    decrease_button.pack(side=tk.LEFT, padx=5, pady=5)

    # 创建一个框架用于文本框和滚动条
    frame = tk.Frame(new_window)
    frame.pack(fill=tk.BOTH, expand=True)

    # 创建不可编辑的文本框
    text_box = tk.Text(frame, wrap=tk.WORD, state=tk.DISABLED, font=("Arial", 12))  # 设置初始字体
    text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # 创建垂直滚动条
    scroll_bar = tk.Scrollbar(frame, command=text_box.yview)
    scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

    # 将滚动条和文本框关联
    text_box.config(yscrollcommand=scroll_bar.set)

    # 填充内容
    text_box.config(state=tk.NORMAL)
    text_box.insert(tk.END, content)
    text_box.config(state=tk.DISABLED)

    # 当窗口关闭时从字典中移除
    def on_close():
        open_papers.remove(paper_id)
        new_window.destroy()

    new_window.protocol("WM_DELETE_WINDOW", on_close)

if __name__ == '__main__':
    # 创建主窗口
    root = tk.Tk()
    root.title("EZ听说")
    root.geometry("600x400")

    # 创建一个框架用于状态标签和状态显示区域
    status_frame = tk.Frame(root)
    status_frame.pack(fill=tk.X, padx=10, pady=5)  # 设置边距

    # 创建状态标签
    status_label = tk.Label(status_frame, text="状态", width=10, anchor="w")
    status_label.pack(side=tk.LEFT)

    # 创建状态显示区域
    status_box = tk.Text(status_frame, height=1, width=30, state=tk.DISABLED)
    status_box.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # 创建置顶复选框
    topmost_var = tk.BooleanVar()
    topmost_var.set(False)  # 设置初始状态
    topmost_checkbutton = tk.Checkbutton(status_frame, text="置顶", variable=topmost_var, command=toggle_topmost)
    topmost_checkbutton.pack(side=tk.LEFT)  # 将复选框添加到状态框架中
    
    # 创建一个框架用于日志文本框和滚动条
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)

    # 创建不可编辑的文本框用于显示日志内容
    text_box_log = tk.Text(frame, height=10, width=50, state=tk.DISABLED)
    text_box_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 2))

    # 创建垂直滚动条
    scroll_bar = tk.Scrollbar(frame, command=text_box_log.yview)
    scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

    # 将滚动条和文本框关联
    text_box_log.config(yscrollcommand=scroll_bar.set)

    # 启动程序状态检查线程
    program_name = 'ETSShell.exe'
    threading.Thread(target=check_program_status, args=(program_name,), daemon=True).start()

    # 运行主事件循环
    root.mainloop()
