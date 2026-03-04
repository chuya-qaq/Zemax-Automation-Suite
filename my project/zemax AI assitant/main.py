# main.py
import tkinter as tk
# ✅ 更加简洁，利用了 __init__.py 的封装
from ui_core import ZemaxSmartApp 

if __name__ == "__main__":
    root = tk.Tk()
    app = ZemaxSmartApp(root)
    root.mainloop()