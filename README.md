# Zemax 自动化设计套件 （Zemax-Automation-Suit） - 开发说明文档

这是一个基于 **Python 3.8+** 和 **Zemax ZOSAPI** 开发的工程化光学辅助设计套件。本项目旨在通过模块化的工具箱（主要为公差分析，自动降敏两个部分），帮助光学设计者提升效率。

> **⚠️ 开发状态说明：**
> 时间有限，本项目目前处于**半成品阶段**。公差分析部分已打通，但是降敏部分不成熟，仍在尝试与验证阶段。

---

## 🚀 已实现的核心功能

### 1. 🖥️ 工程化 GUI 交互平台

* **前后端解耦：** 使用 `ui_core`（Tkinter）与 `zemax_core`（ZOSAPI 逻辑）分离的架构，相比传统的单脚本工具，具有更好的稳定性。
* **多页面管理：** 内置“智能诊断”与“自动降敏”两个核心功能模块，支持实时日志输出与进度监控。
* **线程安全：** 耗时任务（如 Monte Carlo 模拟或本地优化）在后台线程运行，确保 GUI 界面不卡死。

### 2. 🤖 智能标签诊断 (Smart Labeler)

* **数据驱动分析：** 自动解析 Zemax 公差报告中的 `Worst Offenders`。
* **专家规则库：** 结合 Seidel 系数与光线追迹数据（AOI、光线高度），自动为光学表面标注功能标签，并根据标签下处方。
* **处方建议：** 针对不同标签提供具体的优化修改方向以及全系统微扰的方式进行调整。
  
### 3. 🛠️ 自动化工具集成

* **一键优化准备：** 快速将玻璃设为 `Substitute` 模式，或将有数据的非零非球面系数设为变量。
* **高精度报告：** 一键导出包含 LDE 数据、Seidel 系数分布。
---

## 📂 项目结构

```text
my project/
├── zemax AI assistant/       # 主程序目录
│   ├── main.py               # 程序入口
│   ├── ui_core/              # UI 逻辑层
│   │   ├── app.py            # 主框架与页面调度
│   │   └── page_analysis.py  # 诊断页面逻辑
│   └── zemax_core/           # ZOSAPI 核心层
│       ├── controller.py     # 核心控制器，管理连接与分发
│       └── services_ai.py    # AI 启发式优化算法实现
└── zemax tool/               # 独立工具集（原型验证脚本）
    └── zemax tool.py         # 包含 SmartLabeler 与 Evolution 核心算法

```

---

## 🛠️ 安装与运行建议

1. **环境准备：** - Python 3.8+
* Ansys Zemax OpticStudio (建议 2024 R1 及以上版本)


2. **依赖库：**
```bash
pip install pythonnet pandas

```


3. **启动：**
直接运行 `main.py` 启动集成界面，或运行 `zemax tool.py` 调试核心功能。

---

## 📝 开发者备注（半成品局限性）

* **连接路径：** `zemax_core/config.py` 中的 `ZEMAX_PATH` 需根据本地安装路径手动配置。
* **算法鲁棒性：**目前的降敏部分不够成熟。
* **UI 细节：** 部分高级配置项（如多阶段演化参数）尚未完全开放至 GUI，需在 `services_ai.py` 中手动修改。
