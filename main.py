"""程序入口"""

import os
import sys

import pythoncom

# 确保项目根目录在 sys.path 中，使 gui/ 包内模块能正确导入根级模块
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后：exe 所在目录
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 开发环境：脚本所在目录
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import tkinter as tk

from gui import VRChatAutoChat


def main():
    pythoncom.CoInitialize()  # 初始化主线程 COM，确保 winsdk 等组件的 COM 环境在整个生命周期内稳定
    root = tk.Tk()
    VRChatAutoChat(root)
    root.mainloop()


if __name__ == "__main__":
    main()
