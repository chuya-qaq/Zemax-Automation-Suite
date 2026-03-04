# zemax_core/config.py

# Zemax 安装路径
ZEMAX_PATH = r"D:\Program Files\Ansys Zemax OpticStudio 2024 R1.00"

# Q-Grade 标准库 (单位：分 Arcmin)
# 注意：TETX/TIRX 在代码中会被除以 60 转换为度
Q_GRADES = {
    "Q1": {"TIND": 0.0001, "TFRN": 0.5, "TTHI": 0.010,  "TIRR": 0.10, "TIRX": 0.17, "TABB": 0.01, "TETX": 0.17, "TSDX": 0.001},
    "Q2": {"TIND": 0.0003, "TFRN": 1.0, "TTHI": 0.010,  "TIRR": 0.10, "TIRX": 0.30, "TABB": 0.03, "TETX": 0.30, "TSDX": 0.003},
    "Q3": {"TIND": 0.0005, "TFRN": 1.0, "TTHI": 0.0125, "TIRR": 0.25, "TIRX": 0.50, "TABB": 0.05, "TETX": 0.50, "TSDX": 0.005},
    "Q4": {"TIND": 0.0008, "TFRN": 2.0, "TTHI": 0.025,  "TIRR": 0.25, "TIRX": 0.80, "TABB": 0.08, "TETX": 0.80, "TSDX": 0.008},
    "Q5": {"TIND": 0.0010, "TFRN": 2.0, "TTHI": 0.0375, "TIRR": 0.50, "TIRX": 1.00, "TABB": 0.10, "TETX": 1.00, "TSDX": 0.010}
}