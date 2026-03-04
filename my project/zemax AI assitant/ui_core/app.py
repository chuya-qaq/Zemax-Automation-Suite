# ui_core/app.py
import tkinter as tk
from tkinter import messagebox
import queue
import threading
import time
from zemax_core.controller import ZemaxController
from .theme import init_styles, COLORS, FONTS
from .page_analysis import PageAnalysis
from .page_desens import PageDesens

class ZemaxSmartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zemax AI OpticStudio Copilot 2026")
        self.root.geometry("1280x900")
        
        self.zx = ZemaxController()
        self.msg_queue = queue.Queue()
        self.is_connected = False
        self.is_running = False
        
        init_styles()
        self._build_main_structure()
        self._poll_messages()

    def _build_main_structure(self):
        sidebar = tk.Frame(self.root, bg="#111827", width=80) 
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)
        
        self.status_frame = tk.Frame(sidebar, bg="#111827", height=100, cursor="hand2")
        self.status_frame.pack(fill=tk.X, pady=(20, 10))
        
        self.light_canvas = tk.Canvas(self.status_frame, width=40, height=40, bg="#111827", highlightthickness=0)
        self.light_canvas.pack(pady=5)
        self.light_id = self.light_canvas.create_oval(5, 5, 35, 35, fill=COLORS["neon_off"], outline="")
        
        self.lbl_connect_text = tk.Label(self.status_frame, text="点击连接", fg="#6B7280", bg="#111827", font=("微软雅黑", 8))
        self.lbl_connect_text.pack()

        self.status_frame.bind("<Button-1>", self.cmd_connect)
        self.light_canvas.bind("<Button-1>", self.cmd_connect)
        self.lbl_connect_text.bind("<Button-1>", self.cmd_connect)
        
        self.nav_btns = {}
        tk.Frame(sidebar, bg="#374151", height=1).pack(fill=tk.X, padx=10, pady=10)
        
        for k, i, l in [("analysis", "📊", "诊断"), ("desens", "🛡️", "降敏")]:
            f = tk.Frame(sidebar, bg="#111827", height=70, cursor="hand2")
            f.pack(fill=tk.X); f.pack_propagate(False)
            
            bar = tk.Frame(f, bg=COLORS["neon_on"], width=4)
            icon = tk.Label(f, text=i, font=("Segoe UI Emoji", 16), bg="#111827", fg="#9CA3AF")
            icon.pack(pady=(12,0))
            txt = tk.Label(f, text=l, font=("微软雅黑", 8), bg="#111827", fg="#9CA3AF")
            txt.pack()
            
            f.bind("<Button-1>", lambda e, key=k: self.switch_page(key))
            icon.bind("<Button-1>", lambda e, key=k: self.switch_page(key))
            txt.bind("<Button-1>", lambda e, key=k: self.switch_page(key))
            self.nav_btns[k] = {"frame": f, "icon": icon, "text": txt, "bar": bar}

        self.main_area = tk.Frame(self.root, bg=COLORS["bg_main"])
        self.main_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        header = tk.Frame(self.main_area, bg="white", height=70, padx=25)
        header.pack(fill=tk.X, side=tk.TOP); header.pack_propagate(False)
        tk.Label(header, text="ZEMAX COPILOT", fg=COLORS["dark"], bg="white", font=FONTS["header"]).pack(side=tk.LEFT, pady=15)
        self.lbl_subtitle = tk.Label(header, text="", fg="#9CA3AF", bg="white", font=FONTS["sub"])
        self.lbl_subtitle.pack(side=tk.LEFT, pady=(22, 15))
        
        self.page_container = tk.Frame(self.main_area, bg=COLORS["bg_main"], padx=20, pady=20)
        self.page_container.pack(fill=tk.BOTH, expand=True)

        self.page_analysis = PageAnalysis(self.page_container, self)
        self.page_desens = PageDesens(self.page_container, self)
        
        self.switch_page("analysis")

    def switch_page(self, key):
        if key == "analysis":
            self.page_desens.pack_forget()
            self.page_analysis.pack(fill=tk.BOTH, expand=True)
            self.lbl_subtitle.config(text="  |  智能公差诊断 (System Diagnostic)")
        elif key == "desens":
            self.page_analysis.pack_forget()
            self.page_desens.pack(fill=tk.BOTH, expand=True)
            self.lbl_subtitle.config(text="  |  AI 自动降敏 (Auto Desensitization)")
        
        for k, btn in self.nav_btns.items():
            active = (k == key)
            fg = "white" if active else "#6B7280"
            bg = "#1F2937" if active else "#111827"
            
            btn["frame"].config(bg=bg)
            btn["icon"].config(fg=fg, bg=bg)
            btn["text"].config(fg=fg, bg=bg)
            if active: btn["bar"].place(x=0, y=0, relheight=1.0)
            else: btn["bar"].place_forget()

    def cmd_connect(self, event=None):
        if self.is_connected: return 
        
        self.light_canvas.itemconfig(self.light_id, fill="#F59E0B") 
        self.lbl_connect_text.config(text="连接中...", fg="#F59E0B")
        self.root.update()
        
        s, m = self.zx.connect()
        if s:
            self.is_connected = True
            self.light_canvas.itemconfig(self.light_id, fill=COLORS["neon_on"], outline="white")
            self.lbl_connect_text.config(text="已连接", fg=COLORS["neon_on"])
            self.page_analysis.btn_sens.config(state=tk.NORMAL)
            self.page_analysis.lbl_prog_1.config(text="就绪", fg=COLORS["success"])
            messagebox.showinfo("Zemax", m)
        else:
            self.light_canvas.itemconfig(self.light_id, fill=COLORS["neon_off"])
            self.lbl_connect_text.config(text="连接失败", fg=COLORS["danger"])
            messagebox.showerror("连接错误", m)

    def start_thread(self, mode):
        if self.is_running: return
        self.is_running = True
        
        # 🟢 【新增】 锁定界面，防止误操作
        if mode == "sensitivity":
            self.page_analysis.lock_ui_for_run()
            
        threading.Thread(target=self._worker, args=(mode,), daemon=True).start()

 

    def _worker(self, mode):
        def cb(p, t): self.msg_queue.put(("p", p, t))
        def log_func(msg): self.msg_queue.put(("log", msg))
        
        try:
            if mode == "sensitivity":
                # ... (sensitivity logic unchanged)
                self.msg_queue.put(("log", "开始系统诊断流程..."))
                
                current_grade = self.page_analysis.cmb_grade.get()
                criterion = self.page_analysis.crit_mode.get() 
                use_mc = self.page_analysis.mc_enable.get()
                
                mc_runs = 0
                if use_mc:
                    try: mc_runs = int(self.page_analysis.mc_runs.get())
                    except: mc_runs = 500
                
                ana_param = self.page_analysis.ana_param.get()
                
                settings = {
                    'grade': current_grade,
                    'criterion': criterion,
                    'mc_runs': mc_runs,
                    'ana_param': ana_param
                }
                
                res = self.zx.run_sensitivity_analysis(cb, log_cb=log_func, grade_settings=settings)
                self.msg_queue.put(("sens", res))
                
            elif mode == "ai":
                
                self.msg_queue.put(("log", "开始 AI 降敏..."))
                strat_id = self.page_desens.selected_strategy.get()
                
                # 🟢 [修改] 获取新的约束参数 (支持范围和不等式)
                constraints = {
                    "effl_min": self.page_desens.effl_min.get().strip(),
                    "effl_max": self.page_desens.effl_max.get().strip(),
                    "wfno_op": self.page_desens.wfno_op.get(),
                    "wfno_val": self.page_desens.wfno_val.get().strip()
                }
                
                res = self.zx.run_ai_heuristic_optimization(strat_id, cb, log_cb=log_func, constraints=constraints)
                self.msg_queue.put(("ai", res))
                
        except Exception as e:
             self.msg_queue.put(("log", f"ERROR: {str(e)}"))
        finally:
            self.msg_queue.put(("done", None))

    # ... (rest unchanged)

    def _poll_messages(self):
        try:
            while True:
                msg_data = self.msg_queue.get_nowait()
                type_ = msg_data[0]
                val = msg_data[1]
                
                if type_ == "p":
                    self.page_analysis.pb_1['value'] = val
                    self.page_desens.pb_2['value'] = val
                    if len(msg_data) > 2:
                        self.page_analysis.lbl_prog_1.config(text=msg_data[2])

                elif type_ == "log":
                    ts = time.strftime("%H:%M:%S")
                    log_entry = f"[{ts}] > {val}\n"
                    self.page_analysis.txt_log_1.insert(tk.END, log_entry)
                    self.page_analysis.txt_log_1.see(tk.END)
                    self.page_desens.txt_log_2.insert(tk.END, log_entry)
                    self.page_desens.txt_log_2.see(tk.END)

                elif type_ == "sens":
                    self.page_analysis.txt_report.delete(1.0, tk.END)
                    self.page_analysis.txt_report.insert(tk.END, val["report"])
                    self.page_analysis.lbl_prog_1.config(text="诊断分析完成", fg=COLORS["success"])
                    self.page_desens.unlock_strategies(val.get("data", [])) 

                elif type_ == "ai":
                    if isinstance(val, dict):
                        self.page_desens.update_result_ui(val)
                    else:
                        self.page_desens.txt_log_2.insert(tk.END, f"\n[RESULT]\n{val}\n")
                    self.page_desens.txt_log_2.insert(tk.END, "[SYSTEM] AI 流程结束。\n")

                elif type_ == "done":
                    self.is_running = False
                    self.page_analysis.txt_log_1.insert(tk.END, "-"*30 + " 任务结束 " + "-"*30 + "\n")
                    # 🟢 【新增】 任务结束，解锁 UI
                    self.page_analysis.unlock_ui_after_run()

        except queue.Empty:
            pass
        self.root.after(100, self._poll_messages)