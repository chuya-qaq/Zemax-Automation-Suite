# zemax_core/services_ai.py
import time
import math
import os
import tempfile
import System # type: ignore

# =============================================================================
# 1. 基础工具函数
# =============================================================================

def _log(cb, msg):
    if cb: cb(msg)

def _safe_close_tool(tool):
    if tool:
        try:
            if hasattr(tool, "Cancel"): tool.Cancel()
            if hasattr(tool, "Close"): tool.Close()
        except: pass

def _force_set_solve(cell, solve_type_enum):
    """[底层核心] 强制设置单元格的求解类型"""
    try:
        solver = cell.CreateSolveType(solve_type_enum)
        if solver is not None:
            cell.SetSolveData(solver)
            return True
    except: pass
    try:
        cell.Solve = solve_type_enum
        return True
    except: pass
    return False

def _unlock_all_semi_diameters(system, ZOSAPI, log_cb=None):
    """[关键] 解锁所有半口径为自动。"""
    try:
        LDE = system.LDE
        ColSD = ZOSAPI.Editors.LDE.SurfaceColumn.SemiDiameter
        SolveType = ZOSAPI.Editors.SolveType
        count = 0
        unlocked_list = []
        for i in range(1, LDE.NumberOfSurfaces): 
            try:
                surf = LDE.GetSurfaceAt(i)
                sd_cell = surf.GetCellAt(ColSD)
                if hasattr(sd_cell, "MakeSolveAutomatic"):
                    sd_cell.MakeSolveAutomatic()
                    count += 1
                    unlocked_list.append(str(i))
                else:
                    _force_set_solve(sd_cell, SolveType.Automatic)
                    count += 1
                    unlocked_list.append(str(i))
            except: pass
        if log_cb and count > 0:
            _log(log_cb, f"   -> [LDE] 已解锁半口径: Surfaces {','.join(unlocked_list)}")
        return count
    except: return 0

def _get_current_mtf(system, ZOSAPI, freq):
    """获取 FFT MTF 平均值"""
    try:
        target_wave_idx = 0
        try:
            num_waves = system.SystemData.Wavelengths.NumberOfWavelengths
            if num_waves < 2:
                target_wave_idx = 1
        except: pass

        analysis = system.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.FftMtf)
        settings = analysis.GetSettings()
        try:
            settings.MaximumFrequency = float(freq)
            settings.Wavelength.SetWavelengthNumber(target_wave_idx)
        except Exception:
            pass 

        analysis.ApplyAndWaitForCompletion()
        results = analysis.GetResults()
        temp_file = os.path.join(tempfile.gettempdir(), "zos_ai_mtf_temp.txt")
        results.GetTextFile(temp_file)
        
        mtf_values = []
        if os.path.exists(temp_file):
            with open(temp_file, 'r', encoding='utf-16-le') as f:
                lines = f.readlines()
            target_freq = float(freq)
            for line in lines:
                parts = line.strip().split()
                if len(parts) < 3: continue
                try:
                    line_freq = float(parts[0])
                    if abs(line_freq - target_freq) < 0.01:
                        val_t = float(parts[1])
                        val_s = float(parts[2])
                        mtf_values.append((val_t + val_s) / 2.0)
                except ValueError:
                    continue 
            
            analysis.Close()
            try: os.remove(temp_file)
            except: pass
            
            if mtf_values:
                return sum(mtf_values) / len(mtf_values)
            else: return 0.0
        else:
            analysis.Close()
            return 0.0
    except Exception as e:
        return 0.0

def _get_seidel_aberrations(system, ZOSAPI, log_cb=None):
    """提取实际的赛德尔像差系数作为基准。"""
    seidel_data = {}
    try:
        analysis = system.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.SeidelCoefficients)
        analysis.ApplyAndWaitForCompletion()
        results = analysis.GetResults()
        temp_file = os.path.join(tempfile.gettempdir(), "zos_ai_seidel_temp.txt")
        results.GetTextFile(temp_file)
        analysis.Close()

        if os.path.exists(temp_file):
            with open(temp_file, 'r', encoding='utf-16-le') as f:
                content = f.read()
            capture = False
            for line in content.splitlines():
                if "Surf" in line and "S1" in line:
                    capture = True
                    continue
                if capture:
                    if line.strip() == "" or "Sum" in line: break
                    parts = line.split()
                    if len(parts) >= 6 and parts[0].isdigit():
                        try:
                            s_idx = int(parts[0])
                            vals = {
                                'SPHA': abs(float(parts[1])),
                                'COMA': abs(float(parts[2])),
                                'ASTI': abs(float(parts[3])),
                                'FCUR': abs(float(parts[4])),
                                'DIST': abs(float(parts[5]))
                            }
                            significant = {k: v for k, v in vals.items() if v > 1e-5}
                            if significant: seidel_data[s_idx] = significant
                        except: pass
            try: os.remove(temp_file)
            except: pass
    except Exception as e:
        if log_cb: _log(log_cb, f"⚠️ [Analysis] Seidel 分析过程异常: {e}")
    return seidel_data

# =============================================================================
# 2. 全局变量配置
# =============================================================================

def _set_all_materials_substitute(system, ZOSAPI, log_cb):
    lde = system.LDE
    num_surf = lde.NumberOfSurfaces
    count = 0
    target_catalog = "CDGM-ZEMAX202409"
    changed_surfs = []
    
    try:
        SolveType = ZOSAPI.Editors.SolveType
        substitute_enum = SolveType.MaterialSubstitute
    except: return 0

    for i in range(1, num_surf - 1):
        try:
            surf = lde.GetSurfaceAt(i)
            mat_name = str(surf.Material).strip().upper()
            if not mat_name or mat_name == "MIRROR": continue

            solve_data = surf.MaterialCell.CreateSolveType(substitute_enum)
            try: solve_data.S_MaterialSubstitute.Catalog = target_catalog
            except: pass
            
            surf.MaterialCell.SetSolveData(solve_data)
            count += 1
            changed_surfs.append(str(i))
        except: pass
    
    if count > 0:
        _log(log_cb, f"   -> [LDE] 玻璃库替换(Substitute)已激活: Surfaces {', '.join(changed_surfs)}")
    return count

def _set_nonzero_aspheric_variables(system, ZOSAPI, log_cb):
    lde = system.LDE
    num_surf = lde.NumberOfSurfaces
    count_vars = 0
    details = []
    
    try:
        SolveType = ZOSAPI.Editors.SolveType
        SurfaceColumn = ZOSAPI.Editors.LDE.SurfaceColumn
    except: return 0

    for i in range(1, num_surf):
        try:
            surf = lde.GetSurfaceAt(i)
            s_type_str = str(surf.Type).strip().upper()
            
            is_asphere = False
            if "EVEN" in s_type_str or s_type_str == "23": is_asphere = True
            elif "Q-TYPE" in s_type_str or "QTYPE" in s_type_str: is_asphere = True
            
            if is_asphere:
                for p_idx in range(1, 19):
                    col_name = f"Par{p_idx}"
                    if not hasattr(SurfaceColumn, col_name): continue
                    try:
                        col_enum = getattr(SurfaceColumn, col_name)
                        cell = surf.GetSurfaceCell(col_enum)
                        try: val = cell.DoubleValue
                        except: val = 0.0
                        if abs(val) > 1e-25:
                            if _force_set_solve(cell, SolveType.Variable):
                                count_vars += 1
                                details.append(f"S{i}.P{p_idx}")
                    except: pass
        except: pass
    
    if count_vars > 0:
        _log(log_cb, f"   -> [LDE] 非球面系数优化已激活: {', '.join(details)}")
    return count_vars

def _setup_global_variables(system, ZOSAPI, log_cb):
    LDE = system.LDE
    try:
        SolveType = ZOSAPI.Editors.SolveType 
        SurfaceColumn = ZOSAPI.Editors.LDE.SurfaceColumn
    except: return

    _log(log_cb, "🔓 [LDE] 正在初始化全局优化配置...")
    try: system.Tools.RemoveAllVariables()
    except: pass
    
    count_geo = 0
    for i in range(1, LDE.NumberOfSurfaces - 1):
        surf = LDE.GetSurfaceAt(i)
        try:
            if abs(surf.Radius) < 1e9: 
                r_cell = surf.GetSurfaceCell(SurfaceColumn.Radius)
                if _force_set_solve(r_cell, SolveType.Variable): count_geo += 1
            if abs(surf.Thickness) > 1e-4:
                t_cell = surf.GetSurfaceCell(SurfaceColumn.Thickness)
                if _force_set_solve(t_cell, SolveType.Variable): count_geo += 1
            if abs(surf.Conic) > 1e-9:
                c_cell = surf.GetSurfaceCell(SurfaceColumn.Conic)
                if _force_set_solve(c_cell, SolveType.Variable): count_geo += 1
        except: pass

    _set_nonzero_aspheric_variables(system, ZOSAPI, log_cb)
    _set_all_materials_substitute(system, ZOSAPI, log_cb)
    _unlock_all_semi_diameters(system, ZOSAPI, log_cb)
    _log(log_cb, f"✅ 全局变量配置完毕: 几何自由度 {count_geo}")

# =============================================================================
# 3. 敏感项提取
# =============================================================================

def _analyze_sensitivity_pathology(ctrl, system, log_cb, top_n=5):
    diagnosis = {}
    offenders = getattr(ctrl, 'last_sensitivity_data', [])
    sens_source = "Controller Cache"
    
    if not offenders:
        try:
            lens_dir = os.path.dirname(system.SystemFile)
            report_path = os.path.join(lens_dir, "AI_Sensitivity_Report.txt")
            if os.path.exists(report_path):
                from .services_ana import _parse_tolerance_report_impl
                _parse_tolerance_report_impl(ctrl, report_path)
                offenders = getattr(ctrl, 'last_sensitivity_data', [])
                sens_source = "File Report"
        except: pass

    _log(log_cb, "🩺 [DIAG] 正在运行实时赛德尔像差分析...")
    real_seidel_map = _get_seidel_aberrations(system, ctrl.ZOSAPI, log_cb)

    count_sens = 0
    if offenders:
        _log(log_cb, f"   -> [Data] 敏感度来源: {sens_source} | 条目: {len(offenders)}")
        for item in offenders:
            s_idx = item.get('surf', 0)
            tol_type = item.get('type', '').upper()
            if s_idx <= 0: continue
            
            if s_idx not in diagnosis:
                if count_sens < top_n:
                    diagnosis[s_idx] = {'types': set(), 'pathology': set(), 'source': 'SENS'}
                    count_sens += 1
                else: continue 
            
            diagnosis[s_idx]['types'].add(tol_type)
            
            path = diagnosis[s_idx]['pathology']
            if 'TIR' in tol_type or 'TED' in tol_type: path.update(['COMA', 'ASTI'])
            elif 'TTH' in tol_type: path.update(['SPHA', 'ASTI'])
            elif 'TRAD' in tol_type: path.update(['SPHA', 'FCUR'])
            elif 'TIND' in tol_type: path.update(['SPHA', 'FCUR'])

    sorted_surfs = []
    for s, vals in real_seidel_map.items():
        weight = sum(vals.values()) 
        sorted_surfs.append((s, weight))
    
    sorted_surfs.sort(key=lambda x: x[1], reverse=True)
    count_seidel = 0
    
    for s_idx, weight in sorted_surfs[:3]:
        if s_idx not in diagnosis:
             diagnosis[s_idx] = {'types': set(), 'pathology': set(), 'source': 'SEIDEL'}
             diagnosis[s_idx]['types'].add("REAL_ABERR")
             count_seidel += 1
        
        vals = real_seidel_map[s_idx]
        d_path = diagnosis[s_idx]['pathology']
        
        if vals.get('SPHA', 0) > 0.05: d_path.add('SPHA')
        if vals.get('COMA', 0) > 0.02: d_path.add('COMA')
        if vals.get('ASTI', 0) > 0.05: d_path.add('ASTI')
        if vals.get('FCUR', 0) > 0.05: d_path.add('FCUR')
        if vals.get('DIST', 0) > 1.0:  d_path.add('DIST')

    if not diagnosis:
        _log(log_cb, "❌ [Error] 未发现敏感项或显著像差，无法确定优化目标。")
        return {}

    _log(log_cb, f"🔍 [DIAG] 混合诊断完成 (Targeting {len(diagnosis)} Surfaces):")
    for s in sorted(diagnosis.keys()):
        d = diagnosis[s]
        if not d['pathology']: d['pathology'].add('SPHA') 
        src = d.get('source', 'MIXED')
        types_str = "+".join(d['types'])
        path_str = "+".join(d['pathology'])
        _log(log_cb, f"   -> [Target] Surf {s} ({src}): 源=[{types_str}] => 策略=[压制 {path_str}]")
    
    return diagnosis

# =============================================================================
# 4. 核心：绝热演化控制器 (替换了原有的小像差互补)
# =============================================================================

class AdiabaticEvolutionController:
    def __init__(self, ctrl, system, ZOSAPI, log_cb, mtf_target_base, diagnosis_map, constraints):
        self.ctrl = ctrl
        self.system = system
        self.ZOSAPI = ZOSAPI
        self.log_cb = log_cb
        self.mfe = self.system.MFE
        self.lde = self.system.LDE
        self.mtf_target_base = mtf_target_base
        self.diagnosis_map = diagnosis_map
        self.constraints = constraints or {}
        
        # 贪婪自适应动力学参数
        self.dynamic_lambda = 0.05   
        self.min_lambda = 0.005      
        self.max_lambda = 0.20       
        self.feather_weight = 0.05   
        self.max_generations = 30
        self.mtf_guard_threshold = 0.10 
        
        # 从 diagnosis_map 动态生成打击源头
        self.dominant_targets = []
        for s_idx, d_info in diagnosis_map.items():
            self.dominant_targets.append({
                "name": f"Target_S{s_idx}", 
                "surf": s_idx, 
                "types": list(d_info.get('pathology', ['SPHA', 'COMA']))
            })
        
        self.evolution_tasks = [] 
        self.current_target_index = 0

    def _insert_op(self, row, op_type, i1=None, i2=None, v1=None, v2=None, v3=None, v4=None, tgt=0.0, wgt=0.0):
        """写入操作数（容错式写入）"""
        op = self.mfe.InsertNewOperandAt(row)
        op.ChangeType(op_type)
        if i1 is not None:
            try: op.GetCellAt(2).IntegerValue = int(i1)
            except System.Exception: pass
        if i2 is not None:
            try: op.GetCellAt(3).IntegerValue = int(i2)
            except System.Exception: pass
        if v1 is not None:
            try: op.GetCellAt(4).DoubleValue = float(v1)
            except System.Exception: pass
        if v2 is not None:
            try: op.GetCellAt(5).DoubleValue = float(v2)
            except System.Exception: pass
        if v3 is not None:
            try: op.GetCellAt(6).DoubleValue = float(v3)
            except System.Exception: pass
        if v4 is not None:
            try: op.GetCellAt(7).DoubleValue = float(v4)
            except System.Exception: pass
            
        try:
            op.Target = float(tgt)
            op.Weight = float(wgt)
        except System.Exception: pass
        
        return op

    def get_true_seidel_values(self):
        return _get_seidel_aberrations(self.system, self.ZOSAPI, self.log_cb)

    def setup_icebreaker_tank(self):
        self.log_cb("🛡️ [防线部署] 构建靶向几何防线与非对称打击阵列...")
        self.evolution_tasks = []
        OpType = self.ZOSAPI.Editors.MFE.MeritOperandType
        current_row = 1 
        
        try:
            last_surf = self.lde.NumberOfSurfaces - 1
            
            # --- 1. 植入用户物理约束 (焦距 / F#) ---
            c_effl_min = self.constraints.get("effl_min", "")
            c_effl_max = self.constraints.get("effl_max", "")
            has_effl = False
            if c_effl_min and float(c_effl_min) > 0:
                row_effl = current_row
                self._insert_op(row_effl, OpType.EFFL, tgt=0.0, wgt=0.0)
                current_row += 1
                self._insert_op(current_row, OpType.OPGT, i1=row_effl, tgt=float(c_effl_min), wgt=1000.0)
                current_row += 1
                has_effl = True
            if c_effl_max and float(c_effl_max) > 0:
                row_effl = current_row
                self._insert_op(row_effl, OpType.EFFL, tgt=0.0, wgt=0.0)
                current_row += 1
                self._insert_op(current_row, OpType.OPLT, i1=row_effl, tgt=float(c_effl_max), wgt=1000.0)
                current_row += 1
                has_effl = True
            
            if not has_effl:
                temp_op = self.mfe.AddOperand()
                temp_op.ChangeType(OpType.EFFL)
                self.mfe.CalculateMeritFunction()
                curr_effl = temp_op.Value
                self.mfe.RemoveOperandAt(temp_op.OperandNumber)
                self._insert_op(current_row, OpType.EFFL, tgt=curr_effl, wgt=1000.0, v1=0.0)
                current_row += 1
            
            c_wfno_val = self.constraints.get("wfno_val", "")
            if c_wfno_val:
                try:
                    val = float(c_wfno_val)
                    row_wfno = current_row
                    self._insert_op(row_wfno, OpType.WFNO, tgt=0.0, wgt=0.0)
                    current_row += 1
                    
                    op_logic = OpType.OPLT
                    if self.constraints.get("wfno_op") == ">": op_logic = OpType.OPGT
                    self._insert_op(current_row, op_logic, i1=row_wfno, tgt=val, wgt=100.0)
                    current_row += 1
                except: pass

            # --- 2. 物理底线 (厚度，防相交) ---
            for i in range(1, last_surf):
                surf = self.lde.GetSurfaceAt(i)
                mat = str(surf.Material).upper().strip()
                if mat and mat != "AIR":
                    self._insert_op(current_row, OpType.CTGT, i1=i, tgt=0.5, wgt=2.0)
                    current_row += 1
                    self._insert_op(current_row, OpType.ETGT, i1=i, tgt=0.5, wgt=2.0)
                    current_row += 1
                else:
                    self._insert_op(current_row, OpType.CTGT, i1=i, tgt=0.1, wgt=1.0)
                    current_row += 1

            # --- 3. 靶向病因打击：AOI 熔断 ---
            for target in self.dominant_targets:
                surf = target["surf"]
                if surf >= last_surf: continue
                row_raid_m = current_row
                self._insert_op(row_raid_m, OpType.RAID, i1=surf, i2=0, v2=1.0, v4=1.0, tgt=0.0, wgt=0.0)
                current_row += 1
                self._insert_op(current_row, OpType.OPLT, i1=row_raid_m, tgt=45.0, wgt=5.0) 
                current_row += 1
                
                row_raid_c = current_row
                self._insert_op(row_raid_c, OpType.RAID, i1=surf, i2=0, v2=1.0, v4=0.0, tgt=0.0, wgt=0.0)
                current_row += 1
                self._insert_op(current_row, OpType.OPLT, i1=row_raid_c, tgt=45.0, wgt=5.0) 
                current_row += 1

            # --- 4. 非对称靶向打击阵列 ---
            true_values = self.get_true_seidel_values()
            op_map = {"SPHA": OpType.SPHA, "COMA": OpType.COMA, "ASTI": OpType.ASTI, "DIST": OpType.DIST, "FCUR": OpType.FCUR}
            
            for target in self.dominant_targets:
                surf = target["surf"]
                if surf not in true_values: continue
                
                task_indices = []
                for t in target["types"]:
                    if t not in op_map: continue
                    real_val = true_values[surf].get(t, 0.0)
                    
                    row_aberr = current_row
                    self._insert_op(row_aberr, op_map[t], i1=surf, i2=0, tgt=0.0, wgt=0.0)
                    current_row += 1
                    
                    row_abso = current_row
                    self._insert_op(row_abso, OpType.ABSO, i1=row_aberr, tgt=0.0, wgt=0.0)
                    current_row += 1
                    
                    row_oplt = current_row
                    op_tgt = self._insert_op(row_oplt, OpType.OPLT, i1=row_abso, tgt=real_val, wgt=0.0)
                    current_row += 1
                    
                    task = {
                        "name": target["name"], 
                        "surf": surf, 
                        "type": t,
                        "abso_row": row_abso,
                        "current_target": real_val, 
                        "op_object": op_tgt
                    }
                    self.evolution_tasks.append(task)
                    task_indices.append(len(self.evolution_tasks) - 1)
                    
                target["task_indices"] = task_indices

            self.mfe.CalculateMeritFunction()
            
            for task in self.evolution_tasks:
                if task["current_target"] <= 1e-9:
                    real_net_val = self.mfe.GetOperandAt(task["abso_row"]).Value
                    task["current_target"] = real_net_val
                    task["op_object"].Target = real_net_val
                self.log_cb(f"   -> 锁定源头: Surf {task['surf']} [{task['type']}] = {task['current_target']:.4f}")
                
            self.log_cb("✅ 几何破冰阵列部署完成！")
            
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self.log_cb(f"❌ 构建模块发生异常: \n{err}")

    def run_quick_focus(self):
        try:
            qf = self.system.Tools.OpenQuickFocus()
            if qf:
                qf.RunAndWaitForCompletion()
                qf.Close()
        except: pass

    def run_quick_optimization(self, cycles=15):
        try:
            tool = self.system.Tools.OpenLocalOptimization()
            if tool:
                tool.Algorithm = 0 # DLS
                tool.NumberOfCycles = cycles
                try: tool.NumberOfCores = 8
                except: pass
                tool.RunAndWaitForCompletion()
                final_mf = tool.CurrentMeritFunction
                tool.Close()
                if final_mf > 1e8 or final_mf < 0: 
                    return False
                return True
        except: 
            return False
        return False

    def respiratory_recovery(self):
        for task in self.evolution_tasks:
            task["op_object"].Weight = self.feather_weight * 0.1 
        self.log_cb("   💨 [深呼吸] 卸载源头压力，全力恢复像质...")
        self.run_quick_focus()
        self.run_quick_optimization(cycles=20)

    def run_evolution(self):
        self.setup_icebreaker_tank()
        
        initial_mtf = _get_current_mtf(self.system, self.ZOSAPI, 30)
        self.log_cb(f"🏁 初始 MTF: {initial_mtf:.4f} | 初始步长: {self.dynamic_lambda:.1%}")
        current_mtf = initial_mtf
        best_mtf = initial_mtf
        success = True
        
        for gen in range(1, self.max_generations + 1):
            if not self.dominant_targets:
                break
                
            target = self.dominant_targets[self.current_target_index]
            self.current_target_index = (self.current_target_index + 1) % len(self.dominant_targets)
            
            self.log_cb(f"\n🌊 [第 {gen} 代] 非对称打击: {target['name']} (当前步长: {self.dynamic_lambda:.1%})")
            
            for t_idx in target.get("task_indices", []):
                task = self.evolution_tasks[t_idx]
                old_tgt = task["current_target"]
                task["safe_target"] = old_tgt 
                
                new_tgt = old_tgt * (1.0 - self.dynamic_lambda)
                task["current_target"] = new_tgt
                
                task["op_object"].Target = new_tgt
                task["op_object"].Weight = self.feather_weight
                self.log_cb(f"   🎯 单压 {task['type']}@{task['surf']}: {old_tgt:.4f} -> {new_tgt:.4f}")

            for i, task in enumerate(self.evolution_tasks):
                if i not in target.get("task_indices", []):
                    task["op_object"].Weight = 0.0

            self.run_quick_focus()
            success = self.run_quick_optimization(cycles=15) 
            self.run_quick_focus()
            
            if not success:
                self.log_cb("❌ [致命熔断] 物理边界被突破，光线无法追迹或系统畸变过大！")
                self.log_cb("⛔ 任务强制安全终止。系统已锁定在崩溃前最后一次有效状态。")
                break 
                
            new_mtf = _get_current_mtf(self.system, self.ZOSAPI, 30)
            drop_rate = (current_mtf - new_mtf) / current_mtf if current_mtf > 0 else 0
            
            if drop_rate < 0.02: 
                self.log_cb(f"   📈 MTF 坚挺 ({new_mtf:.4f})。系统游刃有余！")
                self.dynamic_lambda = min(self.dynamic_lambda * 1.5, self.max_lambda)
                current_mtf = new_mtf
                if new_mtf > best_mtf: best_mtf = new_mtf
                
            elif drop_rate > self.mtf_guard_threshold: 
                self.log_cb(f"   ⚠️ MTF 暴跌至 {new_mtf:.4f} (Drop: {drop_rate:.1%})。触发回滚与深吸气！")
                for t_idx in target.get("task_indices", []):
                    tk_obj = self.evolution_tasks[t_idx]
                    tk_obj["current_target"] = tk_obj["safe_target"]
                    tk_obj["op_object"].Target = tk_obj["safe_target"]
                    
                self.dynamic_lambda = max(self.dynamic_lambda * 0.5, self.min_lambda)
                self.respiratory_recovery() 
                
                recov_mtf = _get_current_mtf(self.system, self.ZOSAPI, 30)
                self.log_cb(f"   ✨ 深呼吸恢复完毕，恢复至: {recov_mtf:.4f}")
                
                current_mtf = max(recov_mtf, new_mtf) 
                if current_mtf > best_mtf: best_mtf = current_mtf
            else:
                self.log_cb(f"   📊 稳步推进，MTF: {new_mtf:.4f} (Drop: {drop_rate:.1%})")
                current_mtf = new_mtf
                if new_mtf > best_mtf: best_mtf = new_mtf
            
        if success:
            self.log_cb("\n✅ 绝热演化破局降敏圆满结束！")
        
        yield_est = 90.0 + (best_mtf - self.mtf_target_base) * 60
        return {
            "nominal_old": f"{self.mtf_target_base:.3f}",
            "nominal_new": f"{best_mtf:.3f}",
            "yield_new": f"{min(yield_est, 99.5):.1f}%",
            "strategy": "Adiabatic Evolution",
            "data": self.diagnosis_map
        }

# =============================================================================
# 5. 主逻辑入口
# =============================================================================

def run_ai_heuristic_optimization_impl(ctrl, strategy_id, progress_cb, log_cb, constraints=None):
    system = ctrl.system
    ZOSAPI = ctrl.ZOSAPI
    
    if constraints is None: constraints = {}
    if not ctrl.is_connected: return "❌ 系统未连接"
    
    # 提取敏感项作为打击目标
    diagnosis_map = _analyze_sensitivity_pathology(ctrl, system, log_cb, top_n=4)
    if not diagnosis_map:
        return "❌ 无法提取敏感项。请先运行 [AI 诊断] 或检查目录下是否有报告文件。"
    
    if strategy_id == 2:
        nominal_mtf = _get_current_mtf(system, ZOSAPI, 30)
        _log(log_cb, f"📊 [Status] 系统原始 MTF (Base): {nominal_mtf:.4f}")

        lens_dir = os.path.dirname(system.SystemFile)
        chk_file = os.path.join(lens_dir, "_AI_Checkpoint_Last.zmx")
        system.SaveAs(chk_file) 

        _log(log_cb, "🔓 [INIT] 正在预解锁系统以恢复光路...")
        _setup_global_variables(system, ZOSAPI, log_cb)
        
        mtf_after_vars = _get_current_mtf(system, ZOSAPI, 30)
        
        if mtf_after_vars < 0.001 and nominal_mtf > 0.01:
            _log(log_cb, "⚠️ [CRITICAL] 警告：变量配置后 MTF 跌落至 0！极可能是玻璃库缺失导致。")
        else:
            diff = mtf_after_vars - nominal_mtf
            icon = "✅" if abs(diff) < 0.05 else "⚠️"
            _log(log_cb, f"{icon} [Check] 变量配置后 MTF: {mtf_after_vars:.4f} (变化: {diff:.4f})")

        try:
            _log(log_cb, "🚀 启动绝热演化 (Adiabatic Evolution) 引擎...")
            evo_controller = AdiabaticEvolutionController(ctrl, system, ZOSAPI, log_cb, nominal_mtf, diagnosis_map, constraints)
            result = evo_controller.run_evolution()
            
            best_mtf = float(result["nominal_new"])
            if best_mtf > nominal_mtf * 0.9:
                _log(log_cb, "   ✅ 降敏成功，系统已重构并保存。")
                system.SaveAs(chk_file) 
            else:
                _log(log_cb, "   ⚠️ 像质损失较大，已自动回滚。")
                system.LoadFile(chk_file, False)
                
            return result
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            _log(log_cb, f"❌ 运行异常: {str(e)}\n{err}")
            system.LoadFile(chk_file, False)
            return "❌ AI 优化异常终止"
        finally:
            try: os.remove(chk_file)
            except: pass
    else:
        return _log(log_cb, f"⚠️ 策略 {strategy_id} 暂不支持。")

def set_materials_substitute_impl(ctrl, log_cb):
    """一键设为替换 (S)"""
    if not ctrl.is_connected: return "❌ 系统未连接"
    count = _set_all_materials_substitute(ctrl.system, ctrl.ZOSAPI, log_cb)
    return f"✅ 已将 {count} 个玻璃表面设为 Substitute。"

def set_aspheric_variables_impl(ctrl, log_cb):
    """设非球面变量"""
    if not ctrl.is_connected: return "❌ 系统未连接"
    count = _set_nonzero_aspheric_variables(ctrl.system, ctrl.ZOSAPI, log_cb)
    return f"✅ 已激活 {count} 个非球面变量。"