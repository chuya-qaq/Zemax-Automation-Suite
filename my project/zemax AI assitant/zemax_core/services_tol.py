# zemax_core/services_tol.py
from .config import Q_GRADES

def _log(cb, msg):
    if cb: cb(msg)

def lock_all_semi_diameters_impl(ctrl, log_cb=None):
    system = ctrl.system
    ZOSAPI = ctrl.ZOSAPI
    
    try:
        if not ctrl.is_connected: return "系统未连接"

        lde = system.LDE
        try:
            ColSD = ZOSAPI.Editors.LDE.SurfaceColumn.SemiDiameter
            SolveFixed = ZOSAPI.Editors.SolveType.Fixed
        except Exception as e:
            _log(log_cb, f"❌ [LDE] API 定义获取失败: {e}")
            return "初始化失败"

        count = 0
        skip_count = 0
        total_surfs = lde.NumberOfSurfaces
        
        _log(log_cb, f"🔍 [LDE] 智能扫描 {total_surfs} 个表面的半口径状态...")
        
        # 🟢【核心修复】范围改为 0 到 total_surfs (不包含 total_surfs 本身)
        # 这样索引就是 0, 1, ..., 24 (共25个)，不会越界访问 index=25
        for i in range(0, total_surfs):
            try:
                surf = lde.GetSurfaceAt(i)
                sd_cell = surf.GetCellAt(ColSD)
                
                # 智能跳过已锁定的
                is_fixed = False
                try:
                    if sd_cell.Solve.Type == SolveFixed:
                        is_fixed = True
                except: pass

                if is_fixed:
                    skip_count += 1
                    continue  

                # 执行锁定
                current_val = surf.SemiDiameter
                sd_cell.MakeSolveFixed()
                surf.SemiDiameter = current_val
                
                count += 1
            except Exception as inner_e:
                # 捕获单行错误，避免整个崩溃
                _log(log_cb, f"⚠️ Surf {i} 固化操作异常: {inner_e}")
        
        try: system.Tools.BasicOptimizers.RunDataCheck()
        except: pass

        if count > 0:
            msg = f"🔒 已强制固化 {count} 个动态口径，智能跳过 {skip_count} 个"
            _log(log_cb, f"[LDE] {msg}")
            return f"✅ 优化完成 ({msg})"
        else:
            msg = f"✅ 所有 {skip_count} 个口径均已锁定"
            _log(log_cb, f"⏭️ [LDE] {msg}")
            return msg

    except Exception as e:
        _log(log_cb, f"❌ [LDE] 口径固化流程崩溃: {e}")
        return f"锁定口径失败: {str(e)}"



def setup_q_grade_tolerances_impl(ctrl, grade, custom_frn, custom_thi, log_cb):
    system = ctrl.system
    ZOSAPI = ctrl.ZOSAPI
    
    if not ctrl.is_connected: return "系统未连接"
    
    target_grade = grade if grade in Q_GRADES else "Q4"
    specs = Q_GRADES[target_grade].copy()
    
    if custom_frn > 0: specs["TFRN"] = custom_frn
    if custom_thi > 0: specs["TTHI"] = custom_thi
    
    _log(log_cb, f"⚙️ [TDE] 初始化公差编辑器，应用标准: {target_grade}")
        
    try:
        tde = system.TDE
        try: tde.DeleteAllRows()
        except: pass 

        lde = system.LDE
        TolTypes = ZOSAPI.Editors.TDE.ToleranceOperandType
        SurfTypes = ZOSAPI.Editors.LDE.SurfaceType 
        
        count = 0
        
        for i in range(1, lde.NumberOfSurfaces): 
            surf = lde.GetSurfaceAt(i)
            if surf.IsStop: continue

            try: 
                mat_name = str(surf.Material).strip().upper()
                prev_surf = lde.GetSurfaceAt(i-1)
                prev_mat = str(prev_surf.Material).strip().upper()
            except: continue
            
            is_glass = (mat_name != "AIR" and mat_name != "")
            is_prev_glass = (prev_mat != "AIR" and prev_mat != "")
            
            is_front_of_lens = is_glass and not is_prev_glass
            is_back_of_lens = not is_glass and is_prev_glass
            is_inside_glass = is_glass and is_prev_glass

            if is_front_of_lens or is_inside_glass:
                op = tde.AddOperand(); op.ChangeType(TolTypes.TTHI); op.Param1 = i; op.Min, op.Max = -specs["TTHI"], specs["TTHI"]
                op = tde.AddOperand(); op.ChangeType(TolTypes.TIND); op.Param1 = i; op.Min, op.Max = -specs["TIND"], specs["TIND"]
                op = tde.AddOperand(); op.ChangeType(TolTypes.TABB); op.Param1 = i; op.Min, op.Max = -specs["TABB"], specs["TABB"]
                op = tde.AddOperand(); op.ChangeType(TolTypes.TFRN); op.Param1 = i; op.Min, op.Max = -specs["TFRN"], specs["TFRN"]
                op = tde.AddOperand(); op.ChangeType(TolTypes.TIRR); op.Param1 = i; op.Min, op.Max = -specs["TIRR"], specs["TIRR"]
                count += 5

                if surf.Type == SurfTypes.Standard:
                    op = tde.AddOperand(); op.ChangeType(TolTypes.TSDX); op.Param1 = i; op.Min, op.Max = -specs["TSDX"], specs["TSDX"]
                    op = tde.AddOperand(); op.ChangeType(TolTypes.TSDY); op.Param1 = i; op.Min, op.Max = -specs["TSDX"], specs["TSDX"]
                    count += 2
                    _log(log_cb, f"  -> Surf {i} (Glass): +Standard Tols")
                else:
                    try:
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TCON); op.Param1 = i; op.Min, op.Max = -0.002, 0.002
                        count += 1
                        _log(log_cb, f"  -> Surf {i} (Glass): +Asphere TCON")
                    except: pass

                if is_front_of_lens:
                    end_surf_index = i
                    found_end = False
                    for k in range(i + 1, lde.NumberOfSurfaces):
                        next_mat = str(lde.GetSurfaceAt(k).Material).strip().upper()
                        if next_mat == "AIR" or next_mat == "":
                            end_surf_index = k; found_end = True; break
                    
                    if found_end:
                        assy_dec = specs["TSDX"] * 2.0 
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TEDX); op.Param1 = i; op.Param2 = end_surf_index; op.Min, op.Max = -assy_dec, assy_dec
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TEDY); op.Param1 = i; op.Param2 = end_surf_index; op.Min, op.Max = -assy_dec, assy_dec
                        
                        assy_tilt_deg = specs["TETX"] / 60.0 
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TETX); op.Param1 = i; op.Param2 = end_surf_index; op.Min, op.Max = -assy_tilt_deg, assy_tilt_deg
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TETY); op.Param1 = i; op.Param2 = end_surf_index; op.Min, op.Max = -assy_tilt_deg, assy_tilt_deg
                        count += 4
                        _log(log_cb, f"  -> Surf {i}: 元件装配 (Range {i}-{end_surf_index}, Tilt={assy_tilt_deg:.4f}°)")

            elif is_back_of_lens:
                op = tde.AddOperand(); op.ChangeType(TolTypes.TFRN); op.Param1 = i; op.Min, op.Max = -specs["TFRN"], specs["TFRN"]
                op = tde.AddOperand(); op.ChangeType(TolTypes.TIRR); op.Param1 = i; op.Min, op.Max = -specs["TIRR"], specs["TIRR"]
                count += 2
                
                if surf.Type != SurfTypes.Standard:
                    try:
                        op = tde.AddOperand(); op.ChangeType(TolTypes.TCON); op.Param1 = i; op.Min, op.Max = -0.002, 0.002
                        count += 1
                        _log(log_cb, f"  -> Surf {i} (Air): +Asphere TCON (无位移/厚度公差)")
                    except: pass
                else:
                    _log(log_cb, f"  -> Surf {i} (Air): +Surface Form (无位移/厚度公差)")

        try:
            op = tde.AddOperand(); op.ChangeType(TolTypes.COMP); op.Param1 = lde.NumberOfSurfaces - 1; op.Min, op.Max = -1.0, 1.0
            count += 1
            _log(log_cb, f"🔧 [COMP] 添加后焦补偿 (Surf {lde.NumberOfSurfaces - 1})")
        except: pass
        
        if count == 0:
            _log(log_cb, "❌ [TDE] 未识别到任何光学表面")
            return "⚠️ 警告: 未识别到透镜，请检查玻璃库数据。"

        return (f"✅ 公差建模完成 (项数:{count})\n"
                f"⚙️ {grade} (Tilt单位已修正为度, 空气面已隔离)")

    except Exception as e:
        _log(log_cb, f"❌ [TDE] 建模崩溃: {e}")
        return f"公差设定失败: {str(e)}"