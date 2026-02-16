# VRChat-OSC-Chat-Tool v1.1

为VRChat的聊天消息制作的OSC程序，它带有UI！

* 欢迎你为程序提出新功能的建议！
* 喜欢本项目记得点个⭐Star噢~

---

## ✨ 最新更新（v1.1 - 2024.02.15）

### 🎉 新增功能
- **蓝牙心率监测** - 支持检测小米手环/手表等支持 HRS（Heart Rate Service）协议的蓝牙设备心率广播
  - 基于 BLE GATT 协议直接读取心率数据
  - 无需第三方APP，直连手环获取实时心率
  - 支持显示设备名称和当前 BPM
  - 心率数据可自动附加到聊天消息中

---

## 📸 预览

<img width="1616" height="953" alt="预览图" src="https://github.com/user-attachments/assets/99c3968e-9229-495c-806b-24a33bb3d911" />

---

## 📞 有 需求/问题 吗？

* 在 Discord 上添加我的好友并给我发消息！
  * 我的名字： **bili_xiaobingge**
* 在 BiliBili 上通过私信给我发消息！
  * https://space.bilibili.com/65102520

---

## 🚀 软件特征

### 核心功能
* ✅ **文本消息** - 支持最多144字符的消息发送
* ✅ **历史记录** - 保存最近20条发送记录
* ✅ **自动发送** - 可设置5-300秒的间隔自动发送

### 智能附加项
* ✅ **当前电脑时间** - 自动附加 UTC+8 格式时间
* ✅ **当前窗口标题** - 获取当前焦点窗口名称
* ✅ **音乐信息** - 使用 SMTC 协议获取歌曲信息（支持网易云音乐等）
* ✅ **挂机检测** - 检测离开电脑的时长，超过阈值自动显示
* ✅ **硬件监测** - CPU、RAM、GPU 占用信息（*不支持AMD显卡*）
* ✅ **心率监测** - 蓝牙直连手环/手表获取实时心率（**NEW!**）

---

## 📦 无需安装

* 只需下载文件运行即可

---

## 🛠️ 本地构建

*（不太推荐，我写的代码嘛...）*

### 环境要求
* Python 3.11+ https://www.python.org/downloads/

### 安装依赖

```bash
pip install python-osc
pip install pywin32
pip install psutil
pip install GPUtil
pip install winsdk
pip install bleak  # 新增：蓝牙心率监测所需
```

### 构建可执行文件

```bash
pip install PyInstaller
# 或使用可视化工具
pip install auto-py-to-exe
```

### 📚 依赖库说明

| 库名 | 作用 |
|:--------:|:--------:|
|tkinter |	GUI界面构建 |
|python-osc |	OSC协议通信 |
|pywin32	| Windows窗口信息获取 |
|psutil |	系统资源监控 |
|GPUtil |	GPU信息获取 |
|winsdk	| Windows媒体控制接口(SMTC) |
|bleak	| 新增 - 蓝牙(BLE)通信，用于心率监测 |

---

### ❤️ 心率监测使用说明

* 支持的设备
* 小米手环 8/9/10
* 小米手表
* 其他支持 HRS (Heart Rate Service) 协议的蓝牙心率设备

### 使用方法
1. 确保你的设备已开启心率广播功能
2. 在程序中勾选"心率检测"选项
3. 程序会自动扫描并连接附近的蓝牙心率设备
4. 连接成功后，心率数据会显示在调试面板中
5. 发送消息时会自动附加 [❤️:XXX BPM] 格式的心率信息

---

### 🔮 未来计划

* [ ] 重构代码？（目前主打一个够用！ 当然如果你想帮忙...）
* [ ] 我想到的好功能都会写进去，你也可以为我提出 issues

---

### 📝 更新日志
#### v1.1 (2024.02.15)

* ✨ 新增蓝牙心率监测功能（基于 HRS 协议）
  
#### v1.0
* 🎉 初始版本发布
* 基础OSC消息发送
* 时间/窗口/音乐/挂机检测
* 硬件监测功能
