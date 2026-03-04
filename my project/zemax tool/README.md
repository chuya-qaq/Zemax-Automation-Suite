# Zemax Seidel & Lens Reporter Pro (AI d)

这是一个基于 Python 和 Zemax ZOSAPI 开发的高效光学系统分析与自动化工具。它旨在为光学设计工程师提供一键式的系统诊断、像差分析及高级降敏优化功能。

## 🚀 核心功能

### 1. 🤖 智能标签诊断 (Smart Labeler)
- **数据驱动：** 自动读取并解析 Zemax 公差报告中的 "Worst Offenders" 数据。
- **专家规则：** 内置光学专家规则库，自动为每个表面贴上功能标签，如：
    - `High_Field_Driver`: 场曲/畸变高敏区
    - `Pupil_Regime_Master`: 球差核心区
    - `Dominant_Source`: 像差主导源头
    - `Potential_Sponge`: 潜力海绵/像差吸收区
- **处方建议：** 为识别出的问题区域提供针对性的设计修改建议。

### 2. 📉 深度像差分析
- **赛德尔系数报告：** 详细提取 S1 (球差)、S2 (彗差)、S3 (像散)、S4 (场曲)、S5 (畸变) 的各面分布及系统总量。
- **MTF 均值分析：** 自动计算 30 lp/mm 下所有视场的 T/S 平均 MTF 值，快速评估系统画质。
- **LDE 结构导出：** 一键导出完整的镜头结构数据表（Radius, Thickness, Glass, Semi-Dia）。

### 3. 🧬 绝热演化降敏 (Adiabatic Evolution)
- **破局策略：** 采用动态步长（$\lambda$）和非对称打击策略，针对特定的像差源头施加“不平衡压力”。
- **自适应优化：** 在优化过程中自动监控 MTF 波动，实现“贪婪加速”或“回滚呼吸”，帮助系统跳出局部极小值，寻找更稳定的全局最优解。

### 4. 🛠️ 自动化工具箱
- **一键 Substitute：** 将系统中所有玻璃材料设为“替代 (S)”状态，准备进行玻璃库优化。
- **非球面变量自动设置：** 自动识别所有非零的非球面系数并将其设为变量。
- **本地优化触发：** 无需切换界面，直接从工具端发起 DLS 本地优化。

## 📦 技术架构
- **语言：** Python 3.x
- **核心库：** ZOSAPI (Ansys Zemax OpticStudio)
- **GUI：** Tkinter
- **依赖：** `pythonnet` (clr), `pandas`, `re`, `threading`

## 📖 使用说明
1. 确保已安装 Ansys Zemax OpticStudio 并能正常启动。
2. 运行 `Export_Seidel_LDE_Report.py`。
3. 点击 **🔗 连接 OpticStudio** 按钮建立交互连接。
4. 加载公差报告文件或直接生成分析报告。
5. 使用高级工具（如绝热演化）前建议先备份原始 Zemax 文件。

---
*本项目由光学设计专家规则驱动，旨在提升设计效率。*
