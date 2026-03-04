# ui_core/theme.py
import tkinter as tk
from tkinter import ttk

COLORS = {
    "primary": "#2563EB", "success": "#059669", "danger": "#DC2626", 
    "warning": "#D97706", "dark": "#1F2937", "text": "#4B5563",
    "disabled": "#9CA3AF", "bg_main": "#F3F4F6", "bg_card": "#FFFFFF",
    "neon_on": "#06B6D4", "neon_off": "#374151"
}

FONTS = {
    "header": ("Impact", 22), "sub": ("微软雅黑", 10),
    "card": ("微软雅黑", 11, "bold"), "body": ("微软雅黑", 9),
    "code": ("Consolas", 10), "big": ("DIN Alternate", 24, "bold")
}

def init_styles():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TProgressbar", thickness=10, background=COLORS["primary"], borderwidth=0)
    style.configure("TNotebook", background=COLORS["bg_main"], borderwidth=0)
    style.configure("TNotebook.Tab", padding=[15, 8], font=FONTS["body"])
    style.map("TNotebook.Tab", background=[("selected", "white")], foreground=[("selected", COLORS["primary"])])
    style.configure("Vertical.TScrollbar", background="#E5E7EB", troughcolor="#F3F4F6", borderwidth=0, arrowsize=10)