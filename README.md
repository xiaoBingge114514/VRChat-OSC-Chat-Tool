# VRChat-OSC-Chat-Tool
为VRChat的聊天消息制作的OSC程序，可以在游戏内发送任何想要在聊天框中常驻的消息，例如：音乐（目前使用SMTC）、离开电脑的时间、电脑性能、目前正在看的窗口

# 有 需求/问题 吗？
在 Discord 上添加我的好友并给我发消息！
我的名字： bili_xiaobingge


# 软件特征
* 文本消息
* 当前的电脑时间
* 目前作为焦点的窗口名称
* 使用 SMTC 协议获取歌曲信息！
* 离开电脑的时间
* CPU、RAM、GPU 占用信息

# 无需安装
* 只需下载文件运行即可

# 本地构建
（不太推荐，我写的代码嘛...)
* 安装python(我使用的3.11) https://www.python.org/downloads/
* 安装所有所需的模块 （pip install {module name}):
  * tkinter
  * python-osc
  * win32gui
  * psutil
  * GPUtil
  * winsdk
* pip install PyInstaller / (pip install auto-py-to-exe 可视化的本地GUI构建)
* 运行构建完的文件

# 依赖库说明：
* tkinter: GUI界面构建
* python-osc: OSC协议通信
* win32gui: Windows窗口信息获取
* psutil: 系统资源监控
* GPUtil: GPU信息获取
* winsdk: Windows媒体控制接口

# 未来：
* 重构代码？ （目前主打一个够用！）
