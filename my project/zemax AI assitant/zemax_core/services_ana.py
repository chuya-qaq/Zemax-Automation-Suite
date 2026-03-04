# zemax_core/services_ana.py
import os
import time
import re

# =================================================================
#  辅助函数
# =================================================================

def _log(cb, msg):
    if cb: cb(msg)

def _safe_close(tool):
    """【防崩溃】强制安全关闭工具"""
    if tool:
        try: tool.Close()
        except: pass

# =================================================================
#  核心逻辑实现
# =================================================================

def run_sensitivity_analysis_impl(ctrl, progress_cb, log_cb, grade_settings):
    system = ctrl.system
    if not ctrl.is_connected: return {"report": "系统未连接", "data": []}
    
    # 重置停止标志
    ctrl.stop_requested = False
    
    tol_tool = None
    temp_path = None
    
    try:
        _log(log_cb, "================ [诊断开始] ================")
        progress_cb(10, "🛡️ 正在强制固化系统...")
        
        # 1. 尝试移除变量
        try: system.Tools.RemoveAllVariables()
        except: pass
        
        # 2. 锁口径
        if hasattr(ctrl, 'lock_all_semi_diameters'):
            ctrl.lock_all_semi_diameters(log_cb)
        
        if ctrl.stop_requested: return {"report": "⛔ 用户已终止任务", "data": []}

        _log(log_cb, "🔒 [LDE] 清除变量并固定求解...")
        
        # 3. 锁半径/厚度/玻璃
        lde = system.LDE
        SolveFixed = ctrl.ZOSAPI.Editors.SolveType.Fixed
        ColRad = ctrl.ZOSAPI.Editors.LDE.SurfaceColumn.Radius
        ColThi = ctrl.ZOSAPI.Editors.LDE.SurfaceColumn.Thickness
        ColMat = ctrl.ZOSAPI.Editors.LDE.SurfaceColumn.Material

        for i in range(1, lde.NumberOfSurfaces + 1):
            try:
                surf = lde.GetSurfaceAt(i)
                # 锁半径
                if surf.GetCellAt(ColRad).SolveType != SolveFixed: 
                    val = surf.Radius
                    surf.GetCellAt(ColRad).MakeSolveFixed()
                    surf.Radius = val
                # 锁厚度
                if surf.GetCellAt(ColThi).SolveType != SolveFixed: 
                    val = surf.Thickness
                    surf.GetCellAt(ColThi).MakeSolveFixed()
                    surf.Thickness = val
                # 锁玻璃
                mat_name = str(surf.Material).strip().upper()
                if mat_name and mat_name != "AIR":
                    if surf.GetCellAt(ColMat).SolveType != SolveFixed:
                        surf.Material = mat_name 
                        surf.GetCellAt(ColMat).MakeSolveFixed()
            except: pass

        # --- 4. 参数准备 ---
        target_grade = grade_settings.get('grade', 'Q4') if grade_settings else 'Q4'
        crit_str = grade_settings.get('criterion', 'MTF')
        mc_runs = grade_settings.get('mc_runs', 0) if grade_settings else 0
        ana_param = grade_settings.get('ana_param', '30') 
        crit_val = grade_settings.get('crit_val', 0.5) if grade_settings else 0.5
        
        param_desc = f"@{ana_param} lp/mm" if crit_str=="MTF" and ana_param else ""
        progress_cb(30, f"⚙️ 标准:{target_grade} | 指标:{crit_str} {param_desc} | MC:{mc_runs}")
        
        # 5. 设置公差
        ctrl.setup_q_grade_tolerances(
            grade=target_grade,
            custom_frn=grade_settings.get('frn', 0.0),
            custom_thi=grade_settings.get('thi', 0.0),
            log_cb=log_cb
        )

        if system.TDE.NumberOfRows < 1:
            return {"report": "❌ 错误：TDE 为空！未识别到公差操作数。", "data": []}

        # 7. 准备报告路径
        lens_path = system.SystemFile
        if not lens_path: return {"report": "❌ 请先保存镜头文件。", "data": []}
        lens_dir = os.path.dirname(lens_path)
        report_name = "AI_Sensitivity_Report.txt"
        temp_path = os.path.join(lens_dir, report_name)
        if os.path.exists(temp_path): os.remove(temp_path)
        
        progress_cb(50, f"🚀 启动 Zemax 内核 ({crit_str})...")
        
        # 8. 打开公差分析工具
        tol_tool = system.Tools.OpenTolerancing() if hasattr(system.Tools, "OpenTolerancing") else system.Tools.OpenTolerance()
        tol_tool.SetupMode = 0 
        
        if crit_str == "RMS":
            tol_tool.Criterion = 3 
            _log(log_cb, "ℹ️ 评价标准: RMS Wavefront")
        else:
            tol_tool.Criterion = 8 
            _log(log_cb, f"ℹ️ 评价标准: Diffraction MTF {param_desc}")
            if ana_param: _log(log_cb, "⚠️ 提示: 分析频率依赖于 Zemax 主界面当前的 MTF 设置")
            
        # 设置蒙特卡洛
        if mc_runs > 0:
            tol_tool.NumberOfRuns = mc_runs
            _log(log_cb, f"🎲 启用蒙特卡洛: {mc_runs} runs")
        else:
            tol_tool.NumberOfRuns = 0
            _log(log_cb, "⏩ 跳过蒙特卡洛 (仅灵敏度)")

        try: tol_tool.OutputFile = report_name 
        except: pass
        
        # 9. 运行 (异步循环，支持取消)
        tol_tool.Run()
        
        is_running = True
        while is_running:
            if ctrl.stop_requested:
                _log(log_cb, "🛑 接到终止指令，正在停止 Zemax 内核...")
                tol_tool.Cancel()
                tol_tool.Close()
                return {"report": "⛔ 分析已由用户手动终止。\n(部分数据可能未生成)", "data": []}
            
            if not tol_tool.IsRunning:
                is_running = False
            
            time.sleep(0.5)
        
        _safe_close(tol_tool); tol_tool = None 
        
        progress_cb(80, "分析完成，正在解析数据...")
        _log(log_cb, "✅ [KERNEL] 计算完成，等待 IO...")
        
        # 11. 等待文件生成
        found_file = False
        for i in range(30): 
            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 1000: 
                found_file = True
                break
            time.sleep(0.5)
        
        if found_file: time.sleep(1.0)
        else: return {"report": "❌ 错误：Zemax 未能生成报告文件 (超时)。", "data": []}

        # 12. 解析报告
        report_text = _parse_tolerance_report_impl(ctrl, temp_path, crit_str, ana_param, crit_val)
        
        count = len(ctrl.last_sensitivity_data)
        _log(log_cb, f"📊 [DATA] 成功提取 {count} 个敏感项")
        _log(log_cb, "================ [诊断结束] ================")

        return {"report": report_text, "data": ctrl.last_sensitivity_data}

    except Exception as e:
        if tol_tool: _safe_close(tol_tool)
        err = f"❌ 运行异常: {str(e)}"
        _log(log_cb, err)
        return {"report": err, "data": []}

# =================================================================
#  报告解析逻辑 (增强修复版)
# =================================================================

def _parse_tolerance_report_impl(ctrl, filepath, crit_str="MTF", ana_param=None, crit_val=None):
    diagnosis_map = {
        "TTHI": ("厚度敏感：光焦度偏差", "高级球差"),
        "TRAD": ("半径敏感：校正负荷", "场曲/球差"),
        "TIND": ("折射率：批次误差", "焦点偏移"),
        "TABB": ("阿贝数：色散波动", "色差"),
        "TIRR": ("不规则度：高频", "高频像差"),
        "TSDX": ("表面偏心：非对称", "像散"),
        "TSDY": ("表面偏心：非对称", "像散"),
        "TEDX": ("元件偏心：核心", "彗差"),
        "TEDY": ("元件偏心：核心", "彗差"),
        "TETX": ("元件倾斜：装配", "像面倾斜"),
        "TETY": ("元件倾斜：装配", "像面倾斜"),
        "TCON": ("非球面：参数误差", "球差")
    }

    try:
        with open(filepath, 'rb') as f: 
            raw = f.read()
            try: content = raw.decode('utf-16-le')
            except: content = raw.decode('mbcs', errors='ignore')
    except: return "❌ 无法读取报告"

    # 1. 自动识别模式
    is_rms = "RMS" in crit_str or "RMS" in content[:1000]
    crit_name = "RMS Wavefront" if is_rms else "Diffraction MTF"
    unit_s = "λ" if is_rms else ""
    
    # 2. 解析 Worst Offenders (敏感度表)
    offenders = []
    capture = False
    start_keys = ["worst offenders", "最坏偏离", "最严重的"]
    # 结束词增加 "Statistics" 以防止读到 Compensator Statistics
    end_keys = ["statistics", "estimated", "nominal", "名义", "compensator"] 
    
    lines = content.splitlines()
    
    # --- 阶段 1: 提取敏感项 ---
    for line in lines:
        s_line = line.strip()
        if not s_line: continue
        line_l = s_line.lower()
        
        if not capture:
            if any(k in line_l for k in start_keys): capture = True
            continue 
            
        if capture:
            if "monte carlo" in line_l or "蒙特卡洛" in line_l: break
            if any(k in line_l for k in end_keys): break
            
            # 兼容 Zemax 新格式 (Value, Criterion, Change 可能会有多个列)
            # 格式: Type Surf [Param] Value Criterion Change
            parts = s_line.split()
            # 确保至少有 3 部分，且第一部分是字母 (Type)
            if len(parts) >= 3 and parts[0][0].isalpha() and "type" not in parts[0].lower():
                try:
                    # Change 始终在最后一列
                    chg = float(parts[-1])
                    t_surf = 0
                    # 尝试找到 Surface Index (通常是第二个参数)
                    for p in parts[1:]: 
                        if p.isdigit(): t_surf = int(p); break
                    
                    offenders.append({'type': parts[0], 'surf': t_surf, 'change': chg, 'abs': abs(chg)})
                except: pass
                
    offenders.sort(key=lambda x: x['abs'], reverse=True)
    ctrl.last_sensitivity_data = offenders

    # --- 阶段 2: 解析统计值 (状态机模式) ---
    nom, est = 0.0, 0.0
    mc_mean = 0.0
    
    in_mc_section = False # 核心修复：状态标志
    
    value_pattern = r"[:：]\s*([-+]?[\d\.]+)"

    for line in lines:
        line_clean = line.strip()
        if not line_clean: continue
        line_l = line_clean.lower()
        
        # 状态切换检测
        if "monte carlo" in line_l or "蒙特卡洛" in line_l:
            in_mc_section = True
        
        # 1. 匹配 Nominal (设计基准) - 只要匹配到就读取，不管在哪里
        if "nominal criterion" in line_l or "nominal mtf" in line_l or "nominal rms" in line_l:
             m = re.search(value_pattern, line_clean)
             if m: nom = float(m.group(1))

        # 2. 匹配 Estimated (RSS 预估值) - 必须排除 "change"
        if ("estimated mtf" in line_l or "estimated criterion" in line_l or "estimated rms" in line_l):
             if "change" not in line_l: # 排除 "Estimated change" 行
                m = re.search(value_pattern, line_clean)
                if m: est = float(m.group(1))

        # 3. 匹配 Mean (蒙特卡洛均值) - 【核心修复】必须在 MC Section 内
        # 避免读取到 "Compensator Statistics" 中的 Mean
        if in_mc_section:
            if (line_l.startswith("mean") or "平均值" in line_l) and "back focus" not in line_l:
                 m = re.search(value_pattern, line_clean)
                 if m: mc_mean = float(m.group(1))
    
    # 逻辑判定: 只有当 MC 均值有效且大于 0 (避免读取到全0) 时才使用，否则使用 RSS 预估
    # 注意：Zemax 报告中，如果没有运行 MC，MC 部分要么不存在，要么为空。
    final_est = mc_mean if (mc_mean > 0.00001) else est
    
    # 4. 解析蒙特卡洛分布
    mc_stats = []
    mc_capture = False
    mc_keys = ["monte carlo", "蒙特卡洛", "蒙特卡罗"]
    
    for line in lines:
        line_l = line.strip().lower()
        if not mc_capture:
            if any(k in line_l for k in mc_keys):
                mc_capture = True
                continue
                
        if mc_capture:
            if "%" in line:
                m = re.search(r"(\d+)\s*%\s*([><]?)\s*([\d\.]+)", line)
                if m: mc_stats.append((m.group(1), m.group(2), m.group(3)))
            if "configuration" in line_l: break

    # 5. 生成报告文本
    out = []
    
    out.append(f"📊 [ 设计基准 ] Nominal {crit_name}: {nom:.4f}{unit_s}")
    if is_rms:
        if nom > 0.25: out.append(f"⚠️ 警告：名义值 > 0.25λ (瑞利判据)，设计余量极低！")
        elif nom > 0.07: out.append(f"ℹ️ 提示：名义值 > 0.07λ (衍射极限)。")
    out.append("="*90)
    
    out.append(f"{'Rank':<5} | {'Type':<6} | {'Surf':<4} | {'Change(Δ)':<10} | {'评级':<8} | {'成因'}")
    out.append("-" * 90)
    
    if not offenders:
        out.append("   (未检测到敏感项，可能是解析未匹配到数据)")
    
    for i, item in enumerate(offenders[:15]):
        level = "🔥" if i<3 else ("⚠️" if item['abs']>(0.05 if is_rms else 0.08) else " ")
        diag = diagnosis_map.get(item['type'], ("常规公差项",""))[0]
        if len(diag) > 15: diag = diag[:15] + "..."
        out.append(f"{level}{i+1:<3} | {item['type']:<6} | {item['surf']:<4} | {item['change']:>10.5f} | {'敏感' if i<3 else ' ':<8} | {diag}")
    
    out.append("="*90)
    
    if is_rms:
        degrad = final_est - nom
        out.append(f"✅ [ 统计结论 ] 预计平均 RMS: {final_est:.4f}{unit_s} (恶化: +{degrad:.4f}{unit_s})")
    else:
        # 使用修正后的 drop 值
        drop = nom - final_est
        # 防止显示 0.0000 (如果解析失败)
        display_est = final_est if final_est > 0.00001 else est
        if display_est < 0.00001: display_est = 0.0 # 实在没有数据
        
        drop = nom - display_est
        out.append(f"✅ [ 统计结论 ] 预计平均 MTF: {display_est:.4f} (跌落: {drop:.4f})")

    # --- 蒙特卡洛结果展示 ---
    if mc_stats:
        out.append("-" * 90)
        out.append("🎲 [ 蒙特卡洛良率预估 (Monte Carlo) ]")
        found_key = False
        for p, s, v in mc_stats:
            if p in ['98', '90', '80', '50', '20', '10', '2']:
                sign = s if s else ("<" if is_rms else ">")
                out.append(f"   • {p}% 产品 {crit_name} {sign} {v}{unit_s}")
                found_key = True
        
        if not found_key:
             for p, s, v in mc_stats[:8]:
                sign = s if s else ("<" if is_rms else ">")
                out.append(f"   • {p}% 产品 {crit_name} {sign} {v}{unit_s}")
    else:
        out.append("-" * 90)
        out.append("ℹ️ 未检测到蒙特卡洛数据 (请确认勾选了 MC 且 Zemax 已生成统计)")
    
    out.append("\n💡 [ AI 综述 ]")
    if is_rms:
        val_90 = next((float(v) for p, s, v in mc_stats if p=='90'), None)
        if val_90 and val_90 > 0.25:
            out.append(f"▶ 生产风险：高。90% 的成品 RMS > 0.25λ (当前: {val_90}{unit_s})，无法满足瑞利判据。")
        elif nom > 0.1:
            out.append(f"▶ 设计基础：较差 ({nom:.4f}{unit_s})。起始波像差较大，公差空间受限。")
        else:
            out.append(f"▶ 状态：良好。公差对波像差影响在可控范围内。")
    else:
        drop = nom - final_est
        if abs(drop) > 1.0: 
             out.append("⚠️ 数据解析异常：MTF 跌落值不合理，请检查 Zemax 报告原件。")
        elif drop < 0.1: out.append("▶ 系统鲁棒性：优秀。")
        elif drop < 0.2: out.append("▶ 系统鲁棒性：中等。")
        else: out.append("⚠️ 系统敏感！存在大像差互补风险。")

    return "\n".join(out)

# zemax_core/services_ana.py 追加到文件末尾

import re

class SmartLabeler:
    """智能贴标签引擎"""
    def __init__(self, system, zosapi, log_cb):
        self.system = system
        self.ZOSAPI = zosapi
        self.log_cb = log_cb
        self.PRESCRIPTIONS = {
            "High_Field_Driver": "⚠️ 场曲/畸变高敏区 -> 需控制主光线角，避免过度弯曲。",
            "Pupil_Regime_Master": "🛑 球差核心区 -> 任何变动将导致全场MTF崩塌，冻结或微调。",
            "Air_Gap_Lever": "📐 空气杠杆 -> 对倾斜敏感，需注意机械定心精度。",
            "Lever_Ratio_High": "🚨 极长空气杠杆 -> 微小倾斜即导致巨大偏心，需最高级装配工艺。",
            "Massive_Reciprocal_Asti": "⚖️ 像散对冲高压区 -> 目标：降敏。尝试拆分透镜或使用高折射率玻璃。",
            "Massive_Reciprocal_Spha": "🔴 巨量球差对冲 -> 影响轴上点锐度。若此处敏感，需严格控制间隔公差。",
            "Massive_Reciprocal_Dist": "🌀 巨量畸变对冲 -> 边缘视场压力极大。建议：优先松动源头面的曲率。",
            "Aberration_Generator": "🌋 像差发生器 -> 系统像差的主要贡献者。",
            "Dominant_Source": "👑 像差源头 (主导面) -> 它是麻烦制造者，优先修改其曲率以减小下游补偿压力。",
            "High_AOI_Hazard": "📉 高入射角危害 (>45°) -> 公差斜率极陡。优化时加入 RAID 控制光线偏折角。",
            "Tilt_Dominant_Exhaust": "📉 偏心敏感 -> 减小入射角 (AOI)，优化光线偏折角。",
            "Potential_Sponge": "🧽 潜力海绵/泄洪区 -> 平面或低像差面，设为变量以吸收残余像差。"
        }

    def parse_worst_offenders(self, text_data):
        offenders = {}
        pattern = re.compile(r"(T[A-Z]{3})\s+(\d+)\s+")
        for line in text_data.strip().split('\n'):
            match = pattern.search(line)
            if match:
                tol_type, surf_idx = match.group(1), int(match.group(2))
                if surf_idx not in offenders: offenders[surf_idx] = {'types': []}
                if tol_type not in offenders[surf_idx]['types']: offenders[surf_idx]['types'].append(tol_type)
        return offenders

    def get_ray_data(self):
        mfe = self.system.MFE
        nsurf = self.system.LDE.NumberOfSurfaces
        ray_data = {} 
        try:
            type_reay = self.ZOSAPI.Editors.MFE.MeritOperandType.REAY
            type_raid = self.ZOSAPI.Editors.MFE.MeritOperandType.RAID
        except: return {}

        for i in range(1, nsurf):
            try:
                y_m = mfe.GetOperandValue(type_reay, i, 0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0)
                aoi_m = mfe.GetOperandValue(type_raid, i, 0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0)
                y_c = mfe.GetOperandValue(type_reay, i, 0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0)
                aoi_c = mfe.GetOperandValue(type_raid, i, 0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0)
                ray_data[i] = {'y_marg': abs(y_m), 'y_chief': abs(y_c), 'aoi_chief': abs(aoi_c), 'aoi_marg': abs(aoi_m)}
            except: pass
        return ray_data

    def get_seidel_data(self):
        seidel_map = {} 
        try:
            analysis = self.system.Analyses.New_Analysis(self.ZOSAPI.Analysis.AnalysisIDM.SeidelCoefficients)
            analysis.ApplyAndWaitForCompletion()
            import tempfile, os
            tmp = os.path.join(tempfile.gettempdir(), "zos_seidel_tmp.txt")
            analysis.GetResults().GetTextFile(tmp)
            if os.path.exists(tmp):
                with open(tmp, 'r', encoding='utf-16-le') as f:
                    for line in f.readlines():
                        parts = line.split()
                        if len(parts) >= 6 and parts[0].isdigit():
                            seidel_map[int(parts[0])] = {
                                'S1': float(parts[1]), 'S3': float(parts[3]), 'S5': float(parts[5]),
                                'Sum_All': sum([abs(float(x)) for x in parts[1:6]])
                            }
                try: os.remove(tmp)
                except: pass
            analysis.Close()
        except: pass
        return seidel_map

    def analyze(self, worst_offenders_text):
        offenders = self.parse_worst_offenders(worst_offenders_text)
        ray_data = self.get_ray_data()
        seidel_data = self.get_seidel_data()
        
        diagnosis_report = []
        all_y_marg = [d.get('y_marg', 0) for d in ray_data.values()]
        all_y_chief = [d.get('y_chief', 0) for d in ray_data.values()]
        max_y_marg = max(all_y_marg) if all_y_marg else 1.0
        max_y_chief = max(all_y_chief) if all_y_chief else 1.0
        
        self.log_cb("🧠 正在应用专家规则库 (Applied Expert Rules)...")
        for i in range(1, self.system.LDE.NumberOfSurfaces):
            labels = []
            r_dat = ray_data.get(i, {})
            s_dat = seidel_data.get(i, {'S1':0, 'S3':0, 'S5':0, 'Sum_All':0})
            is_offender = i in offenders
            
            thickness = self.system.LDE.GetSurfaceAt(i).Thickness
            
            if r_dat.get('aoi_marg', 0) > 45.0 or r_dat.get('aoi_chief', 0) > 45.0: labels.append("High_AOI_Hazard")
            if thickness > 15.0: labels.append("Lever_Ratio_High")
            elif thickness > 5.0: labels.append("Air_Gap_Lever")
            if r_dat.get('y_marg', 0) > 0.7 * max_y_marg: labels.append("Pupil_Regime_Master")
            if r_dat.get('y_chief', 0) > 0.6 * max_y_chief: labels.append("High_Field_Driver")

            if abs(s_dat['S5']) > 5.0:
                labels.append("Massive_Reciprocal_Dist")
                if s_dat['S5'] * seidel_data.get(i+1, {'S5':0})['S5'] < 0 and abs(s_dat['S5']) > 10.0: labels.append("Dominant_Source")

            if abs(s_dat['S1']) > 1.0 and abs(s_dat['S1']) > 2 * abs(s_dat['S3']): labels.append("Massive_Reciprocal_Spha")
            elif abs(s_dat['S3']) > 0.5: labels.append("Massive_Reciprocal_Asti")

            if s_dat['Sum_All'] > 2.0 and "Dominant_Source" not in labels: labels.append("Aberration_Generator")
            if is_offender and any(t in ['TCON', 'TIRR'] for t in offenders[i]['types']): labels.append("Tilt_Dominant_Exhaust")
            if (not is_offender) and (s_dat['Sum_All'] < 0.2 or abs(self.system.LDE.GetSurfaceAt(i).Radius) > 1e8): labels.append("Potential_Sponge")

            if labels:
                diagnosis_report.append(f"Surface {i}: {labels}")
                for lbl in set(labels):
                    if lbl in self.PRESCRIPTIONS: diagnosis_report.append("  -> " + self.PRESCRIPTIONS[lbl])
                diagnosis_report.append("-" * 60)
        
        return "\n".join(diagnosis_report)

def run_smart_diagnosis_impl(ctrl, text_data, log_cb):
    """运行智能诊断"""
    if not ctrl.is_connected: return "❌ 系统未连接"
    labeler = SmartLabeler(ctrl.system, ctrl.ZOSAPI, log_cb)
    return labeler.analyze(text_data)
    
def generate_full_report_impl(ctrl, progress_cb, log_cb):
    """生成全套分析报告"""
    if not ctrl.is_connected: return "❌ 系统未连接"
    system, ZOSAPI = ctrl.system, ctrl.ZOSAPI
    report = []
    _log(log_cb, "正在初始化数据读取...")
    progress_cb(10, "正在读取 LDE...")
    
    try: report.extend(["================================================================",
                        f" 📑 ZEMAX LENS REPORT | EFL: {system.SystemData.EFLE:.2f} | F/#: {system.SystemData.ImageSpaceFNumber:.2f}",
                        "================================================================\n",
                        " [1] 镜头结构数据 (Lens Data)", "-" * 75])
    except: pass

    for i in range(system.LDE.NumberOfSurfaces):
        try:
            surf = system.LDE.GetSurfaceAt(i)
            report.append(f" Surf {i:<3} | Rad: {surf.Radius:<10.4f} | Thi: {surf.Thickness:<10.4f} | Mat: {str(surf.Material):<10}")
        except: pass

    progress_cb(50, "正在读取赛德尔系数...")
    report.extend(["\n [2] 赛德尔像差系数 (Seidel Coefficients)", "-" * 75])
    try:
        from .services_ai import _get_seidel_aberrations
        seidel_map = _get_seidel_aberrations(system, ZOSAPI)
        for s_idx, vals in seidel_map.items():
            report.append(f" Surf {s_idx:<3} | SPHA: {vals.get('SPHA',0):<8.4f} | COMA: {vals.get('COMA',0):<8.4f} | ASTI: {vals.get('ASTI',0):<8.4f}")
    except: pass

    progress_cb(80, "正在获取 MTF...")
    report.extend(["\n [3] MTF 分析报告", "-" * 75])
    try:
        from .services_ai import _get_current_mtf
        mtf_val = _get_current_mtf(system, ZOSAPI, 30)
        report.append(f" 🟢 平均 MTF 值 (30 lp/mm): {mtf_val:.4f}" if mtf_val > 0 else " 🔴 MTF 获取失败或无数据")
    except: pass

    progress_cb(100, "生成完成")
    return "\n".join(report)