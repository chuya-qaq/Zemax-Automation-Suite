"""
Zemax Copilot UI Core Package
=============================
负责应用程序的前端界面构建、样式管理及交互逻辑。
包含主窗口、系统诊断页及 AI 降敏页组件。
"""

import sys
import os

# 包元数据
__version__ = "1.0.0"
__author__ = "Zemax AI Team"

# -----------------------------------------------------------------------------
# 1. Windows 高分屏 (DPI) 适配
# -----------------------------------------------------------------------------
# 这是开发 Tkinter 专业应用的关键步骤。
# 如果不加这段，在高分辨率屏幕(缩放>100%)上，文字和线条会变模糊。
try:
    from ctypes import windll
    # 设置 DPI 感知级别 (1 = System DPI Aware, 2 = Per Monitor DPI Aware)
    # 使用 1 通常最稳定，能防止界面模糊
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    # 非 Windows 系统或调用失败时忽略，不影响程序运行
    pass

# -----------------------------------------------------------------------------
# 2. 暴露核心组件
# -----------------------------------------------------------------------------
# 这样外部只需 `from ui_core import ZemaxSmartApp` 即可，无需深入子模块
from .app import ZemaxSmartApp
from .theme import init_styles

# -----------------------------------------------------------------------------
# 3. 定义导出白名单
# -----------------------------------------------------------------------------
# 当外部执行 `from ui_core import *` 时，只导入以下内容
__all__ = [
    "ZemaxSmartApp",
    "init_styles",
    "__version__"
]