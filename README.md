# **此项目不再维护**

## ~~EZ听说~~: 用于获取 E听说中学 听说测试答案（广东版）
+ 原理: 监听E听说运行日志，获取当前考试状态，并从本地缓存文件中解析试卷答案  
+ 语言: Python 3.x  
+ UI实现: Tkinter  
+ 打包工具: PyInstaller
+ 演示:   


https://github.com/user-attachments/assets/250ee4ac-887c-4a62-adda-dad14b38ec36


+ 使用方法：
1. 前往 [Github 发布页面](https://github.com/victor-egg/EZ-tingshuo/releases)下载最新版本
2. 运行 E听说中心 和 此程序
3. 开始考试,享受 "EZ"

+ 构建:
1. 安装Python包依赖
2. SHELL 运行 `pyinstaller -a --clean main.spec`
3. 输出文件 `dist/EZ.exe`