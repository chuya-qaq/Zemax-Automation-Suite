# ui_core/components.py
import tkinter as tk
from .theme import COLORS, FONTS

def create_card(parent, title):
    card = tk.Frame(parent, bg=COLORS["bg_card"], padx=20, pady=20)
    card.pack(fill=tk.X, pady=(0, 20))
    card.config(highlightbackground="#E5E7EB", highlightthickness=1)
    if title:
        tk.Label(card, text=title, font=FONTS["card"], fg=COLORS["dark"], 
                 bg=COLORS["bg_card"]).pack(anchor="w", pady=(0, 15))
    return card