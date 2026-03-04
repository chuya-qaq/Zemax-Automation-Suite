# ui_core/page_desens.py
import tkinter as tk
from tkinter import ttk, messagebox
from .theme import COLORS, FONTS
from .components import create_card

class PageDesens:
    def __init__(self, parent, app_instance):
        self.parent = parent
        self.app = app_instance
        self.frame = tk.Frame(parent, bg=COLORS["bg_main"])
        self.selected_strategy = tk.IntVar(value=2) # 默认选策略2
        self.strategy_widgets = []
        
        # 🟢 [修改] 约束变量升级
        # EFFL 改为范围 (Min ~ Max)
        self.effl_min = tk.StringVar()
        self.effl_max = tk.StringVar()
        
        # WFNO 改为 不等式 (Operator + Value)
        self.wfno_op = tk.StringVar(value="<") # 默认小于
        self.wfno_val = tk.StringVar()
        
        self._build_ui()

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
    
    def pack_forget(self):
        self.frame.pack_forget()

    def _build_ui(self):
        left_panel = tk.Frame(self.frame, bg=COLORS["bg_main"], width=450)
        left_panel.pack(side=tk.LEFT, fill=tk.Y)
        left_panel.pack_propagate(False)

        # === 1. 策略选择 ===
        c_strat = create_card(left_panel, "1. 降敏策略选择 (AI Strategy)")
        
        self.strategy_widgets = []
        strats = [
            ("结构鲁棒化 (Global)", "重构光焦度分布以降低灵敏度"),
            ("小像差互补 (Balance)", "利用相对表面抵消残余像差"),
            ("公差分级松弛 (Grade)", "针对敏感面分配精密公差"),
            ("主动补偿 (Compensator)", "定义最佳装配补偿器算子"),
            ("玻璃库替换 (Glass)", "匹配更稳健的材料折射率斜率")
        ]
        
        for idx, (title, sub) in enumerate(strats):
            f = tk.Frame(c_strat, bg="#F9FAFB", padx=10, pady=8, height=60)
            f.pack(fill=tk.X, pady=4); f.pack_propagate(False)
            
            # 只有策略2是完全实装的，其他置灰
            state = tk.NORMAL if idx == 1 else tk.DISABLED
            
            rb = tk.Radiobutton(f, variable=self.selected_strategy, value=idx+1, 
                                bg="#F9FAFB", command=self._update_strat_ui, state=state)
            rb.place(x=0, y=10)
            
            lbl_title = tk.Label(f, text=title, font=("微软雅黑", 9, "bold"), 
                                 bg="#F9FAFB", fg=COLORS["disabled"] if state==tk.DISABLED else COLORS["dark"])
            lbl_title.place(x=25, y=5)
            
            tk.Label(f, text=sub, font=("微软雅黑", 8), bg="#F9FAFB", fg="#D1D5DB").place(x=25, y=28)
            
            self.strategy_widgets.append({
                "frame": f,
                "rb": rb,
                "title": lbl_title
            })
        
        # === 1.5 核心参数约束 (新增) ===
        c_constr = create_card(left_panel, "1.5 核心参数约束 (Constraints)")
        
        # 🟢 [修改] EFL 行 (Min ~ Max)
        f_effl = tk.Frame(c_constr, bg="white")
        f_effl.pack(fill=tk.X, pady=5)
        tk.Label(f_effl, text="焦距范围 (EFFL):", width=14, anchor="w", bg="white", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        # Min 输入框
        tk.Entry(f_effl, textvariable=self.effl_min, width=6, bg="#F9FAFB", highlightthickness=1).pack(side=tk.LEFT)
        tk.Label(f_effl, text=" ~ ", bg="white").pack(side=tk.LEFT)
        # Max 输入框
        tk.Entry(f_effl, textvariable=self.effl_max, width=6, bg="#F9FAFB", highlightthickness=1).pack(side=tk.LEFT)
        
        tk.Label(f_effl, text=" mm", fg=COLORS["disabled"], bg="white", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=(5,0))

        # 🟢 [修改] F/# 行 (Inequality)
        f_wfno = tk.Frame(c_constr, bg="white")
        f_wfno.pack(fill=tk.X, pady=5)
        tk.Label(f_wfno, text="工作 F/# (WFNO):", width=14, anchor="w", bg="white", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        # 运算符下拉框 (<, >, =)
        self.cb_wfno_op = ttk.Combobox(f_wfno, textvariable=self.wfno_op, values=["<", ">", "="], width=3, state="readonly")
        self.cb_wfno_op.pack(side=tk.LEFT, padx=(0,5))
        
        # 值输入框
        tk.Entry(f_wfno, textvariable=self.wfno_val, width=8, bg="#F9FAFB", highlightthickness=1).pack(side=tk.LEFT)
        tk.Label(f_wfno, text=" (留空锁当前)", fg=COLORS["disabled"], bg="white", font=("微软雅黑", 8)).pack(side=tk.LEFT)

        # === 2. 执行 ===
        c_exec = create_card(left_panel, "2. AI 优化执行")
        self.pb_2 = ttk.Progressbar(c_exec, mode='determinate', style="TProgressbar")
        self.pb_2.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_ai = tk.Button(c_exec, text="⚡ 启动 AI 降敏引擎", bg=COLORS["success"], fg="white", 
                                font=("微软雅黑", 10, "bold"), relief="flat", pady=12, state=tk.DISABLED, 
                                command=lambda: self.app.start_thread("ai"))
        self.btn_ai.pack(fill=tk.X)

        # === 右侧结果展示 (保持不变) ===
        right_panel = tk.Frame(self.frame, bg=COLORS["bg_main"])
        right_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        nb_page2 = ttk.Notebook(right_panel, style="TNotebook")
        nb_page2.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        t_report = tk.Frame(nb_page2, bg="white")
        nb_page2.add(t_report, text="  📈 优化决策报告  ")

        c_res = tk.Frame(t_report, bg="white", padx=20, pady=20)
        c_res.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(c_res, text="📊 降敏代价-收益评估", font=("微软雅黑", 11, "bold"), 
                 bg="white", fg=COLORS["dark"]).pack(anchor="w", pady=(0, 20))
        
        box = tk.Frame(c_res, bg="white")
        box.pack(fill=tk.X, pady=10)

        f_cost = tk.Frame(box, bg="#FFF1F2", padx=20, pady=25, highlightthickness=1, highlightbackground="#FECACA")
        f_cost.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 10))
        tk.Label(f_cost, text="Nominal MTF (设计名义值)", bg="#FFF1F2", fg=COLORS["danger"], font=("微软雅黑", 10)).pack(anchor="w")
        self.v_nom = tk.Label(f_cost, text="--", font=("Arial", 28, "bold"), bg="#FFF1F2", fg=COLORS["danger"])
        self.v_nom.pack(pady=15)
        tk.Label(f_cost, text="⚠️ 此项通常会略微下降", bg="#FFF1F2", fg="#F87171", font=("微软雅黑", 8)).pack()

        f_gain = tk.Frame(box, bg="#ECFDF5", padx=20, pady=25, highlightthickness=1, highlightbackground="#A7F3D0")
        f_gain.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(10, 0))
        tk.Label(f_gain, text="Production Yield (量产良率)", bg="#ECFDF5", fg=COLORS["success"], font=("微软雅黑", 10)).pack(anchor="w")
        self.v_yld = tk.Label(f_gain, text="--", font=("Arial", 28, "bold"), bg="#ECFDF5", fg=COLORS["success"])
        self.v_yld.pack(pady=15)
        tk.Label(f_gain, text="🚀 核心提升指标", bg="#ECFDF5", fg="#34D399", font=("微软雅黑", 8, "bold")).pack()

        f_btns = tk.Frame(c_res, bg="white")
        f_btns.pack(fill=tk.X, side=tk.BOTTOM, pady=20)
        
        self.btn_accept = tk.Button(f_btns, text="✅ 采纳并保存此方案", bg=COLORS["success"], fg="white", 
                                    font=("微软雅黑", 9, "bold"), relief="flat", padx=30, pady=10, 
                                    state=tk.DISABLED, command=lambda: messagebox.showinfo("成功", "方案已保存至镜头文件。"))
        self.btn_accept.pack(side=tk.RIGHT)
        
        self.btn_undo = tk.Button(f_btns, text="↩️ 撤销优化并还原", bg="#E5E7EB", fg=COLORS["text"], 
                                  font=("微软雅黑", 9), relief="flat", padx=20, pady=10, 
                                  state=tk.DISABLED, command=lambda: messagebox.showinfo("还原", "已回滚至原始设计状态。"))
        self.btn_undo.pack(side=tk.LEFT)

        t_log = tk.Frame(nb_page2, bg="#1E1E1E")
        nb_page2.add(t_log, text="  💻 调试日志  ")
        self.txt_log_2 = tk.Text(t_log, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4", relief="flat", padx=15, pady=15)
        self.txt_log_2.pack(fill=tk.BOTH, expand=True)
        
        self._update_strat_ui()

    def _update_strat_ui(self):
        idx = self.selected_strategy.get()
        for i, w_dict in enumerate(self.strategy_widgets):
            bg_color = "#EFF6FF" if i+1==idx else "white"
            w_dict["frame"].config(bg=bg_color)
        self.btn_ai.config(state=tk.NORMAL)

    def unlock_strategies(self, data):
        for i, widget_dict in enumerate(self.strategy_widgets):
            if i == 1:
                widget_dict["frame"].config(bg="white", cursor="hand2")
                widget_dict["rb"].config(state=tk.NORMAL)
                widget_dict["title"].config(fg=COLORS["dark"])

    def update_result_ui(self, res):
        try:
            nom_old = float(res.get("nominal_old", 0))
            nom_new = float(res.get("nominal_new", 0))
            delta_nom = nom_old - nom_new
            yield_str = str(res.get("yield_new", "0%"))
            yield_val = float(yield_str.replace('%', ''))

            if delta_nom < 0.12:
                c_nom_bg, c_nom_fg, c_nom_bd, msg_nom = "#ECFDF5", "#059669", "#A7F3D0", "✅ 性能牺牲极小"
            elif delta_nom < 0.28:
                c_nom_bg, c_nom_fg, c_nom_bd, msg_nom = "#FFFBEB", "#D97706", "#FDE68A", "⚠️ 画质有一定折损"
            else:
                c_nom_bg, c_nom_fg, c_nom_bd, msg_nom = "#FEF2F2", "#DC2626", "#FECACA", "🚨 性能大幅跌落"

            if nom_new < 0.4:
                c_nom_fg = "#DC2626"
                msg_nom = "❌ 最终画质过低"

            if yield_val > 75:
                c_yld_bg, c_yld_fg, c_yld_bd, msg_yld = "#ECFDF5", "#059669", "#A7F3D0", "🚀 理想量产状态"
            elif yield_val > 40:
                c_yld_bg, c_yld_fg, c_yld_bd, msg_yld = "#FFFBEB", "#D97706", "#FDE68A", "⚠️ 建议小规模试产"
            else:
                c_yld_bg, c_yld_fg, c_yld_bd, msg_yld = "#FEF2F2", "#DC2626", "#FECACA", "❌ 生产即亏损"

            self.v_nom.master.config(bg=c_nom_bg, highlightbackground=c_nom_bd)
            self.v_nom.config(text=str(nom_new), fg=c_nom_fg, bg=c_nom_bg)
            
            self.v_yld.master.config(bg=c_yld_bg, highlightbackground=c_yld_bd)
            self.v_yld.config(text=yield_str, fg=c_yld_fg, bg=c_yld_bg)

            self.btn_accept.config(state=tk.NORMAL)
            self.btn_undo.config(state=tk.NORMAL)

        except Exception as e:
            self.txt_log_2.insert(tk.END, f"[ERROR] UI渲染失败: {e}\n")