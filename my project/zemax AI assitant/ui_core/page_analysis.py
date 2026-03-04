# ui_core/page_analysis.py
import tkinter as tk
from tkinter import ttk, filedialog
from .theme import COLORS, FONTS
from .components import create_card

class PageAnalysis:
    def __init__(self, parent, app_instance):
        self.parent = parent
        self.app = app_instance 
        self.frame = tk.Frame(parent, bg=COLORS["bg_main"])
        
        # --- 状态变量 ---
        self.analysis_mode = tk.IntVar(value=1)
        self.custom_entries = []
        self.save_path = tk.StringVar(value="C:/Project/Lens_Analysis_V1.zmx")
        
        self.crit_mode = tk.StringVar(value="MTF")
        self.ana_param = tk.StringVar(value="30")
        self.mc_enable = tk.BooleanVar(value=False)
        self.mc_runs = tk.StringVar(value="500")
        
        self.grade_info = {
            "Q1": {"title": "Q1 极高精度", "desc": "光刻/干涉仪 | 成本极高", "color": "#7F1D1D", "bg": "#FEF2F2"},
            "Q2": {"title": "Q2 精密级",   "desc": "高端单反/显微 | 成本高",   "color": "#B91C1C", "bg": "#FEF2F2"},
            "Q3": {"title": "Q3 标准级",   "desc": "手机/车载 | 工业标准",     "color": "#D97706", "bg": "#FFFBEB"},
            "Q4": {"title": "Q4 商业级",   "desc": "监控/消费电子 | 易量产",   "color": "#059669", "bg": "#ECFDF5"},
            "Q5": {"title": "Q5 常规级",   "desc": "玩具/教学 | 成本最低",     "color": "#6B7280", "bg": "#F3F4F6"}
        }

        self._build_ui()

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)
        
    def pack_forget(self):
        self.frame.pack_forget()

    def _build_ui(self):
        left_panel = tk.Frame(self.frame, bg=COLORS["bg_main"], width=450)
        left_panel.pack(side=tk.LEFT, fill=tk.Y)
        left_panel.pack_propagate(False)

        canvas = tk.Canvas(left_panel, bg=COLORS["bg_main"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_panel, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
        scrollable_frame = tk.Frame(canvas, bg=COLORS["bg_main"])

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=420)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- 卡片 1-3 ---
        c1 = create_card(scrollable_frame, "1. 工程归档 (Project Save)")
        f_save = tk.Frame(c1, bg="white")
        f_save.pack(fill=tk.X)
        tk.Entry(f_save, textvariable=self.save_path, bg="#F9FAFB", relief="flat", highlightthickness=1).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        tk.Button(f_save, text="📂", bg="#E5E7EB", relief="flat", width=4, command=self._browse_file).pack(side=tk.RIGHT, padx=(5,0))

        c2 = create_card(scrollable_frame, None)
        rb1_f = tk.Frame(c2, bg="white")
        rb1_f.pack(fill=tk.X, pady=(0, 10))
        tk.Radiobutton(rb1_f, text="2. 标准公差等级 (Q-Grade Mode)", variable=self.analysis_mode, value=1, 
                       font=FONTS["card"], bg="white", command=self.toggle_mode).pack(side=tk.LEFT)
        self.cmb_grade = ttk.Combobox(c2, values=list(self.grade_info.keys()), state="readonly")
        self.cmb_grade.current(3) 
        self.cmb_grade.pack(fill=tk.X, pady=(0, 5))
        self.cmb_grade.bind("<<ComboboxSelected>>", self._on_grade_change)
        self.f_tip = tk.Frame(c2, padx=10, pady=8) 
        self.f_tip.pack(fill=tk.X)
        self.lbl_tip_t = tk.Label(self.f_tip, text="", font=("微软雅黑", 9, "bold")); self.lbl_tip_t.pack(anchor="w")
        self.lbl_tip_d = tk.Label(self.f_tip, text="", font=("微软雅黑", 8)); self.lbl_tip_d.pack(anchor="w")
        self._on_grade_change(None) 

        c3 = create_card(scrollable_frame, None)
        rb2_f = tk.Frame(c3, bg="white")
        rb2_f.pack(fill=tk.X, pady=(0, 10))
        tk.Radiobutton(rb2_f, text="3. 自定义公差矩阵 (Custom Matrix)", variable=self.analysis_mode, value=2, 
                       font=FONTS["card"], bg="white", command=self.toggle_mode).pack(side=tk.LEFT)
        f_matrix = tk.Frame(c3, bg="white"); f_matrix.pack(fill=tk.X)
        
        def add_matrix_row(parent_w, items, start_row, label_bg="white", entry_bg="#F9FAFB"):
            for i, (label_text, default_val) in enumerate(items):
                r = start_row + (i // 2)
                c = (i % 2) * 2
                lbl = tk.Label(parent_w, text=label_text, bg=label_bg, fg=COLORS["text"], font=("微软雅黑", 8))
                lbl.grid(row=r, column=c, sticky="e", pady=4, padx=(0, 5))
                ent = tk.Entry(parent_w, width=8, bg=entry_bg, relief="flat", highlightthickness=1, justify="center")
                ent.insert(0, default_val)
                ent.grid(row=r, column=c+1, sticky="w", pady=4, padx=(0, 10))
                self.custom_entries.append((lbl, ent)) 

        basic_p = [("TIND (折射率)", "0.0005"), ("TTHI (厚度)", "0.050"), ("TFRN (半径)", "0.500"), ("TIRR (不规则)", "0.200")]
        f_matrix.columnconfigure(0, weight=1); f_matrix.columnconfigure(2, weight=1)
        add_matrix_row(f_matrix, basic_p, 0)

        self.f_adv_custom = tk.Frame(c3, bg="#F5F3FF", padx=8, pady=8, highlightbackground="#DDD6FE", highlightthickness=1)
        self.f_adv_custom.pack(fill=tk.X, pady=(10, 0))
        self.lbl_adv_title = tk.Label(self.f_adv_custom, text="🚀 AI 高级参数", font=("微软雅黑", 8, "bold"), fg="#7C3AED", bg="#F5F3FF")
        self.lbl_adv_title.pack(anchor="w")
        f_adv_grid = tk.Frame(self.f_adv_custom, bg="#F5F3FF"); f_adv_grid.pack(fill=tk.X)
        adv_p = [("TEDX/Y (偏心)", "0.050"), ("TETX/Y (倾斜)", "0.050"), ("TCON (非球面)", "0.100")]
        f_adv_grid.columnconfigure(0, weight=1); f_adv_grid.columnconfigure(2, weight=1)
        add_matrix_row(f_adv_grid, adv_p, 0, label_bg="#F5F3FF", entry_bg="white")

        # --- 卡片 4: 设置 ---
        c_set = create_card(scrollable_frame, "4. 分析与验证设置 (Settings)")
        f_crit = tk.Frame(c_set, bg="white"); f_crit.pack(fill=tk.X, pady=(0, 10))
        tk.Label(f_crit, text="评价指标:", bg="white", fg=COLORS["text"], font=("微软雅黑", 9, "bold")).pack(side=tk.LEFT, anchor="n", pady=4)
        f_radios = tk.Frame(f_crit, bg="white"); f_radios.pack(side=tk.LEFT, padx=15)
        self.rb_mtf = tk.Radiobutton(f_radios, text="Diffraction MTF (终检)", variable=self.crit_mode, value="MTF", bg="white", command=self._on_crit_change)
        self.rb_mtf.pack(anchor="w")
        self.rb_rms = tk.Radiobutton(f_radios, text="RMS Wavefront (研发)", variable=self.crit_mode, value="RMS", bg="white", command=self._on_crit_change)
        self.rb_rms.pack(anchor="w", pady=(2,0))
        
        self.f_param = tk.Frame(f_crit, bg="white", padx=10); self.f_param.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        self.lbl_param = tk.Label(self.f_param, text="频率 (lp/mm):", bg="white", fg=COLORS["text"], font=("微软雅黑", 8)); self.lbl_param.pack(anchor="w")
        self.ent_param = tk.Entry(self.f_param, textvariable=self.ana_param, width=8, bg="#F9FAFB", highlightthickness=1, justify="center"); self.ent_param.pack(anchor="w", pady=(2,0), ipady=2)

        f_mc = tk.Frame(c_set, bg="#F0FDF4", padx=10, pady=8, highlightbackground="#BBF7D0", highlightthickness=1)
        f_mc.pack(fill=tk.X)
        
        # 🟢 【核心修改】 保存 checkbox 对象到 self.chk_mc
        self.chk_mc = tk.Checkbutton(f_mc, text="启用蒙特卡洛模拟 (Yield Check)", variable=self.mc_enable, bg="#F0FDF4", activebackground="#F0FDF4", font=("微软雅黑", 9, "bold"), command=self._toggle_mc_input)
        self.chk_mc.pack(side=tk.LEFT)
        
        self.lbl_runs = tk.Label(f_mc, text="Runs:", bg="#F0FDF4", fg=COLORS["disabled"]); self.lbl_runs.pack(side=tk.LEFT, padx=(15, 5))
        self.ent_runs = tk.Entry(f_mc, textvariable=self.mc_runs, width=6, state=tk.DISABLED); self.ent_runs.pack(side=tk.LEFT)

        # --- 卡片 5: 执行按钮区 ---
        c_exec = tk.Frame(scrollable_frame, bg=COLORS["bg_main"], pady=20)
        c_exec.pack(fill=tk.X)
        self.pb_1 = ttk.Progressbar(c_exec, mode='determinate', style="TProgressbar")
        self.pb_1.pack(fill=tk.X, pady=(0, 10))
        
        f_btns = tk.Frame(c_exec, bg=COLORS["bg_main"]); f_btns.pack(fill=tk.X)
        
        self.btn_sens = tk.Button(f_btns, text="🚀 开始系统诊断", bg=COLORS["primary"], fg="white", 
                                  font=("微软雅黑", 11, "bold"), relief="flat", pady=15, state=tk.DISABLED, 
                                  command=lambda: self.app.start_thread("sensitivity"))
        self.btn_sens.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.btn_stop = tk.Button(f_btns, text="🛑 终止", bg=COLORS["danger"], fg="white",
                                  font=("微软雅黑", 11, "bold"), relief="flat", pady=15, state=tk.DISABLED,
                                  command=self._cmd_stop)
        self.btn_stop.pack(side=tk.RIGHT, padx=(5, 0))

        self.lbl_prog_1 = tk.Label(c_exec, text="等待指令...", bg=COLORS["bg_main"], fg=COLORS["disabled"], font=("微软雅黑", 8))
        self.lbl_prog_1.pack(pady=5)

        self.toggle_mode()
        self._on_crit_change()

        right_panel = tk.Frame(self.frame, bg=COLORS["bg_main"])
        right_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        nb = ttk.Notebook(right_panel, style="TNotebook")
        nb.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        t1 = tk.Frame(nb, bg="white"); nb.add(t1, text="  📄 分析报告  ")
        self.txt_report = tk.Text(t1, font=FONTS["code"], relief="flat", padx=20, pady=20)
        self.txt_report.pack(fill=tk.BOTH, expand=True)
        t2 = tk.Frame(nb, bg="#1E1E1E"); nb.add(t2, text="  💻 调试日志  ")
        self.txt_log_1 = tk.Text(t2, font=("Consolas", 9), bg="#1E1E1E", fg="#4EC9B0", relief="flat", padx=10, pady=10)
        self.txt_log_1.pack(fill=tk.BOTH, expand=True)

    def _cmd_stop(self):
        self.app.zx.stop_analysis()
        self.lbl_prog_1.config(text="正在终止进程...", fg=COLORS["danger"])
        self.btn_stop.config(state=tk.DISABLED)

    # 🟢 【新增】 锁定 UI：禁止修改 MC、参数等
    def lock_ui_for_run(self):
        self.btn_sens.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        # 锁定 MC 复选框和输入框
        self.chk_mc.config(state=tk.DISABLED)
        self.ent_runs.config(state=tk.DISABLED) 
        # 锁定其他参数
        self.cmb_grade.config(state=tk.DISABLED)
        self.ent_param.config(state=tk.DISABLED)
        self.rb_mtf.config(state=tk.DISABLED)
        self.rb_rms.config(state=tk.DISABLED)
        # 锁定自定义矩阵
        for _, ent in self.custom_entries: ent.config(state=tk.DISABLED)

    # 🟢 【新增】 解锁 UI：恢复状态
    def unlock_ui_after_run(self):
        self.btn_sens.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.chk_mc.config(state=tk.NORMAL)
        # 恢复 MC 输入框（依赖当前勾选状态）
        self._toggle_mc_input() 
        # 恢复其他参数逻辑
        self.rb_mtf.config(state=tk.NORMAL)
        self.rb_rms.config(state=tk.NORMAL)
        self._on_crit_change() # 恢复 param 框状态
        self.toggle_mode()     # 恢复 grade/custom 状态

    def _on_crit_change(self):
        mode = self.crit_mode.get()
        if mode == "MTF":
            self.f_param.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
            self.lbl_param.config(text="频率 (lp/mm):")
            self.ana_param.set("30")
        else:
            self.f_param.pack_forget()
            self.ana_param.set("")

    def _toggle_mc_input(self):
        if self.mc_enable.get():
            self.ent_runs.config(state=tk.NORMAL, bg="white")
            self.lbl_runs.config(fg=COLORS["text"])
        else:
            self.ent_runs.config(state=tk.DISABLED, bg="#F3F4F6")
            self.lbl_runs.config(fg=COLORS["disabled"])

    def _browse_file(self):
        f = filedialog.asksaveasfilename(defaultextension=".zmx", filetypes=[("Zemax", "*.zmx")])
        if f: self.save_path.set(f)

    def _on_grade_change(self, event):
        raw = self.cmb_grade.get().split(" ")[0]
        info = self.grade_info.get(raw, self.grade_info["Q4"])
        self.f_tip.config(bg=info["bg"])
        self.lbl_tip_t.config(text=info["title"], fg=info["color"], bg=info["bg"])
        self.lbl_tip_d.config(text=info["desc"], fg=info["color"], bg=info["bg"])

    def toggle_mode(self):
        mode = self.analysis_mode.get()
        if mode == 1:
            self.cmb_grade.config(state="readonly")
            self.lbl_tip_t.config(fg=COLORS["text"])
            self.lbl_tip_d.config(fg=COLORS["text"])
            for lbl, ent in self.custom_entries:
                lbl.config(fg=COLORS["disabled"])
                ent.config(state=tk.DISABLED, bg=COLORS["bg_main"])
            self.lbl_adv_title.config(fg=COLORS["disabled"])
            self.f_adv_custom.config(highlightbackground="#E5E7EB")
        else:
            self.cmb_grade.config(state=tk.DISABLED)
            self.lbl_tip_t.config(text="等级模式已停用", fg=COLORS["disabled"])
            self.lbl_tip_d.config(text="当前正在手动控制公差细节", fg=COLORS["disabled"])
            for lbl, ent in self.custom_entries:
                lbl.config(fg=COLORS["text"])
                ent.config(state=tk.NORMAL, bg="white")
            self.lbl_adv_title.config(fg="#7C3AED")
            self.f_adv_custom.config(highlightbackground="#DDD6FE")

    def _load_tol_file(self):
        file_path = filedialog.askopenfilename(title="选择公差报告", filetypes=[("Text Files", "*.txt")])
        if not file_path: return
        
        content = ""
        for enc in ['utf-16', 'utf-8', 'gbk']:
            try:
                with open(file_path, 'r', encoding=enc) as f: content = f.read()
                break
            except UnicodeError: continue
        
        extracted = []
        capturing = False
        for line in content.splitlines():
            if "Worst offenders" in line or "最坏偏离" in line:
                capturing = True; extracted.append(line); continue
            if capturing:
                if not line.strip() and len(extracted)>5: break
                if "Estimated Performance" in line or "Monte Carlo" in line: break
                if "-------" not in line: extracted.append(line)
                
        if extracted:
            self.txt_tol_preview.delete(1.0, tk.END)
            self.txt_tol_preview.insert(tk.END, "\n".join(extracted))
            self.app.msg_queue.put(("log", "✅ 已成功提取 Worst Offenders 数据。"))