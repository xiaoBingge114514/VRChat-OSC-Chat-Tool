"""程序入口"""

import tkinter as tk

from gui import VRChatAutoChat


def main():
    root = tk.Tk()
    VRChatAutoChat(root)
    root.mainloop()


if __name__ == "__main__":
    main()
