import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import threading
import os
import sys
import clr
import tempfile
import time
from datetime import datetime
import re

# =================================================================
#  核心逻辑：光学智能标签识别器 (Smart Labeler)
# =================================================================
class SmartLabeler:
    def __init__(self, system, zosapi, log_cb):
        self.system = system
        self.ZOSAPI = zosapi
        self.log_cb = log_cb
        
        # 标签定义字典 (已根据专家意见扩展)
        self.PRESCRIPTIONS = {
            "High_Field_Driver": "⚠️ 场曲/畸变高敏区 -> 需控制主光线角，避免过度弯曲。",
            "Pupil_Regime_Master": "🛑 球差核心区 -> 任何变动将导致全场MTF崩塌，冻结或微调。",
            
            "Air_Gap_Lever": "📐 空气杠杆 -> 对倾斜(Tilt)敏感，需注意机械定心精度。",
            "Lever_Ratio_High": "🚨 极长空气杠杆 (High Lever) -> 微小倾斜即导致巨大偏心，需最高级装配工艺。",
            
            "Massive_Reciprocal_Asti": "⚖️ 像散对冲高压区 -> 目标：降敏。尝试拆分透镜或使用高折射率玻璃。",
            "Massive_Reciprocal_Spha": "🔴 巨量球差对冲 -> 影响轴上点锐度。若此处敏感，需严格控制间隔公差。",
            "Massive_Reciprocal_Dist": "🌀 巨量畸变对冲 -> 边缘视场压力极大。建议：优先松动源头面的曲率。",
            
            "Aberration_Generator": "🌋 像差发生器 -> 系统像差的主要贡献者。",
            "Dominant_Source": "👑 像差源头 (主导面) -> 它是麻烦制造者，优先修改其曲率以减小下游补偿压力。",
            
            "High_AOI_Hazard": "📉 高入射角危害 (>45°) -> 公差斜率极陡。处方：优化时加入 RAID 控制光线偏折角。",
            "Tilt_Dominant_Exhaust": "📉 偏心敏感 -> 减小入射角 (AOI)，优化光线偏折角。",
            
            "Thickness_Insensitive": "🛡️ 厚度钝感 -> 可作为“吸气期”调节焦面的理想缓冲面。",
            "Potential_Sponge": "🧽 潜力海绵/泄洪区 -> 平面或低像差面，设为变量以吸收残余像差。"
        }

    def parse_worst_offenders(self, text_data):
        """ 解析 Worst Offenders 文本数据 """
        offenders = {}
        pattern = re.compile(r"(T[A-Z]{3})\s+(\d+)\s+")
        lines = text_data.strip().split('\n')
        for line in lines:
            match = pattern.search(line)
            if match:
                tol_type = match.group(1)
                surf_idx = int(match.group(2))
                if surf_idx not in offenders:
                    offenders[surf_idx] = {'types': [], 'raw': []}
                if tol_type not in offenders[surf_idx]['types']:
                    offenders[surf_idx]['types'].append(tol_type)
        return offenders

    def get_ray_data(self):
        """ 使用 MFE 获取光线数据 (修复版) """
        self.log_cb("⚡ 正在执行光线追迹 (AOI/Height check)...")
        mfe = self.system.MFE
        nsurf = self.system.LDE.NumberOfSurfaces
        ray_data = {} 
        
        try:
            OpType = self.ZOSAPI.Editors.MFE.MeritOperandType
            type_reay = OpType.REAY
            type_raid = OpType.RAID
        except Exception:
            return {}

        for i in range(1, nsurf):
            try:
                # 边缘光线 (Hy=0, Py=1)
                y_m = mfe.GetOperandValue(type_reay, i, 0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0)
                aoi_m = mfe.GetOperandValue(type_raid, i, 0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0)
                # 主光线 (Hy=1, Py=0)
                y_c = mfe.GetOperandValue(type_reay, i, 0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0)
                aoi_c = mfe.GetOperandValue(type_raid, i, 0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0)
                
                ray_data[i] = {'y_marg': abs(y_m), 'y_chief': abs(y_c), 'aoi_chief': abs(aoi_c), 'aoi_marg': abs(aoi_m)}
            except Exception: pass
        return ray_data

    def get_seidel_data(self):
        """ 获取详细的赛德尔系数 (S1, S3, S5 分离) """
        self.log_cb("📉 读取赛德尔系数分布...")
        seidel_map = {} 
        
        analysis = self.system.Analyses.New_Analysis(self.ZOSAPI.Analysis.AnalysisIDM.SeidelCoefficients)
        analysis.ApplyAndWaitForCompletion()
        results = analysis.GetResults()
        tmp = tempfile.mktemp(suffix=".txt")
        results.GetTextFile(tmp)
        
        if os.path.exists(tmp):
            with open(tmp, 'r', encoding='utf-16-le') as f: lines = f.readlines()
            for line in lines:
                parts = line.split()
                if len(parts) >= 6 and parts[0].isdigit():
                    idx = int(parts[0])
                    try:
                        # 假设列: Surf, S1, S2, S3, S4, S5
                        s1 = float(parts[1])
                        s3 = float(parts[3])
                        s5 = float(parts[5]) # 提取畸变
                        seidel_map[idx] = {
                            'S1': s1, 'S3': s3, 'S5': s5,
                            'Sum_All': sum([abs(float(x)) for x in parts[1:6]])
                        }
                    except: pass
            try: os.remove(tmp)
            except: pass
        analysis.Close()
        return seidel_map

    def analyze(self, worst_offenders_text):
        """ >>> 执行核心诊断流程 (专家校准版) <<< """
        offenders = self.parse_worst_offenders(worst_offenders_text)
        ray_data = self.get_ray_data()
        seidel_data = self.get_seidel_data()
        
        diagnosis_report = []
        
        all_y_marg = [d['y_marg'] for d in ray_data.values()]
        all_y_chief = [d['y_chief'] for d in ray_data.values()]
        max_y_marg = max(all_y_marg) if all_y_marg else 1.0
        max_y_chief = max(all_y_chief) if all_y_chief else 1.0
        
        self.log_cb("🧠 正在应用专家规则库 (Applied Expert Rules)...")
        
        nsurf = self.system.LDE.NumberOfSurfaces
        
        for i in range(1, nsurf):
            labels = []
            
            r_dat = ray_data.get(i, {})
            s_dat = seidel_data.get(i, {'S1':0, 'S3':0, 'S5':0, 'Sum_All':0})
            is_offender = i in offenders
            offender_types = offenders.get(i, {}).get('types', [])
            
            y_m = r_dat.get('y_marg', 0)
            y_c = r_dat.get('y_chief', 0)
            aoi_m = r_dat.get('aoi_marg', 0)
            aoi_c = r_dat.get('aoi_chief', 0)
            thickness = self.system.LDE.GetSurfaceAt(i).Thickness
            radius = self.system.LDE.GetSurfaceAt(i).Radius
            
            # --- A. 物理/结构 规则 ---
            
            # 1. 入射角危害 (新增)
            if aoi_m > 45.0 or aoi_c > 45.0:
                labels.append("High_AOI_Hazard")

            # 2. 空气杠杆 (分级)
            if thickness > 15.0: # 极长间隔 (如 Surface 22)
                labels.append("Lever_Ratio_High")
            elif thickness > 5.0:
                labels.append("Air_Gap_Lever")

            # 3. 瞳面/场面定义
            if y_m > 0.7 * max_y_marg: labels.append("Pupil_Regime_Master")
            if y_c > 0.6 * max_y_chief: labels.append("High_Field_Driver")

            # --- B. 像差特质 规则 (精细化) ---
            
            abs_s1 = abs(s_dat['S1'])
            abs_s3 = abs(s_dat['S3'])
            abs_s5 = abs(s_dat['S5'])

            # 1. 畸变对冲 (Surface 3 & 4 类型)
            if abs_s5 > 5.0: # 畸变系数通常较大
                labels.append("Massive_Reciprocal_Dist")
                # 如果正畸变极大，且下一面负畸变极大，标记当前面为源头
                next_s_dat = seidel_data.get(i+1, {'S5':0})
                if s_dat['S5'] * next_s_dat['S5'] < 0 and abs(s_dat['S5']) > 10.0:
                     labels.append("Dominant_Source")

            # 2. 球差 vs 像散 判别 (Surface 7 & 8 类型)
            # 如果 S1 显著大于 S3，或者是 S3 很小而 S1 较大
            if abs_s1 > 1.0 and abs_s1 > 2 * abs_s3:
                labels.append("Massive_Reciprocal_Spha")
            # 只有当 S3 确实占主导时，才贴像散标签
            elif abs_s3 > 0.5:
                labels.append("Massive_Reciprocal_Asti")

            # 3. 总量检查
            if s_dat['Sum_All'] > 2.0 and "Dominant_Source" not in labels:
                labels.append("Aberration_Generator")

            # --- C. 公差行为 & 角色潜力 ---
            
            if is_offender and ('TCON' in offender_types or 'TIRR' in offender_types):
                labels.append("Tilt_Dominant_Exhaust")
            
            # 潜力海绵 (平面或低像差面)
            # 特别判定：如果是平面 (Radius > 1e8) 且非敏感面
            is_planar = abs(radius) > 1e8
            if (not is_offender) and (s_dat['Sum_All'] < 0.2):
                labels.append("Potential_Sponge")
            elif is_planar and (not is_offender):
                labels.append("Potential_Sponge")

            # --- 生成报告 ---
            if labels:
                line_header = f"Surface {i}: {labels}"
                diagnosis_report.append(line_header)
                unique_prescriptions = set()
                for lbl in labels:
                    if lbl in self.PRESCRIPTIONS:
                        unique_prescriptions.add("  -> " + self.PRESCRIPTIONS[lbl])
                diagnosis_report.extend(list(unique_prescriptions))
                diagnosis_report.append("-" * 60)
        
        return "\n".join(diagnosis_report)
# =================================================================
#  后端逻辑：Zemax 数据提取引擎 (带日志版)
# =================================================================
class ZemaxDataEngine:
    def __init__(self):
        self.app = None
        self.system = None
        self.ZOSAPI = None
        self.is_connected = False
        self.zemax_path = r"D:\Program Files\Ansys Zemax OpticStudio 2024 R1.00"

    def connect(self):
        try:
            if not os.path.exists(self.zemax_path):
                possible_paths = [
                    r"C:\Program Files\Ansys Zemax OpticStudio 2024 R1.00",
                    r"C:\Program Files\Zemax OpticStudio",
                    r"D:\Program Files\Zemax OpticStudio",
                ]
                for p in possible_paths:
                    if os.path.exists(p):
                        self.zemax_path = p
                        break
            
            # 加载 DLL
            net_helper = os.path.join(self.zemax_path, "ZOSAPI_NetHelper.dll")
            if not os.path.exists(net_helper):
                return False, "未找到 ZOSAPI_NetHelper.dll"

            sys.path.append(self.zemax_path)
            if self.zemax_path not in os.environ['PATH']:
                os.environ['PATH'] = self.zemax_path + os.pathsep + os.environ['PATH']
            
            clr.AddReference(net_helper)
            import ZOSAPI_NetHelper
            # 注意：某些版本可能需要捕获 Initialize 的异常
            try:
                ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(self.zemax_path)
            except Exception:
                pass # 如果已经初始化过可能会报错，忽略
            
            clr.AddReference(os.path.join(self.zemax_path, "ZOSAPI.dll"))
            import ZOSAPI
            self.ZOSAPI = ZOSAPI 

            conn = ZOSAPI.ZOSAPI_Connection()
            self.app = conn.ConnectAsExtension(0)
            
            if self.app and self.app.PrimarySystem:
                self.system = self.app.PrimarySystem
                self.is_connected = True
                return True, f"已连接 | ID: {self.system.SystemID}"
            return False, "连接失败"
        except Exception as e:
            return False, str(e)
        
    # =================================================================
    #  【新增】 读取并解析 Zemax 公差报告文件
    # =================================================================
    def load_tolerance_file(self, file_path, log_cb):
        if not os.path.exists(file_path):
            log_cb("❌ 文件不存在")
            return None

        log_cb(f"📂 正在读取文件: {os.path.basename(file_path)}...")
        
        content = ""
        # Zemax 默认生成的文本报告通常是 UTF-16 LE，但也可能是 UTF-8
        encodings = ['utf-16', 'utf-8', 'gbk', 'latin-1']
        
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    content = f.read()
                break # 读取成功
            except UnicodeError:
                continue
        
        if not content:
            log_cb("❌ 无法识别文件编码，读取失败。")
            return None

        # --- 解析逻辑：提取 "Worst offenders" 区域 ---
        lines = content.splitlines()
        extracted_lines = []
        capturing = False
        found_header = False
        
        # 关键词：Zemax 报告中的标准标题
        start_marker = "Worst offenders"
        # 结束标志：可以是空行，也可以是下一个大标题（如 "Estimated Performance"）
        end_markers = ["Estimated Performance", "Monte Carlo Analysis", "Statistics"]

        for line in lines:
            clean_line = line.strip()
            
            # 1. 检测开始
            if start_marker.lower() in clean_line.lower():
                capturing = True
                found_header = True
                extracted_lines.append(line) # 保留标题行
                continue
            
            if capturing:
                # 2. 检测结束 (遇到空行或下一个大标题)
                if not clean_line: 
                    # 有时候 worst offenders 下面紧接着空行，需要多读几行以防万一
                    # 这里简单的逻辑：如果已经提取了数据且遇到连续空行，则停止
                    if len(extracted_lines) > 5: 
                        break
                    else:
                        continue # 跳过头部空行
                
                # 检查是否到了下一个段落
                for end_m in end_markers:
                    if end_m.lower() in clean_line.lower():
                        capturing = False
                        break
                
                if not capturing: break
                
                # 3. 收集数据行 (排除纯虚线分割线)
                if "-------" not in clean_line:
                    extracted_lines.append(line)

        if found_header and len(extracted_lines) > 1:
            log_cb(f"✅ 成功提取 Worst Offenders 数据 ({len(extracted_lines)} 行)")
            return "\n".join(extracted_lines)
        else:
            log_cb("⚠️ 未在文件中找到 'Worst offenders' 数据块。请确认这是公差分析文本报告。")
            return None

    def set_all_materials_substitute(self, log_cb):
        """将所有有效表面的玻璃设置为 Substitute (S) 状态"""
        if not self.is_connected:
            log_cb("❌ 错误：系统未连接")
            return False

        lde = self.system.LDE
        num_surf = lde.NumberOfSurfaces
        count = 0
        target_catalog = "CDGM-ZEMAX202409" 
        
        log_cb(f"正在设置玻璃替换 (S) 模式，目标库: {target_catalog} ...")

        try:
            SolveType = self.ZOSAPI.Editors.SolveType
            substitute_enum = SolveType.MaterialSubstitute
        except AttributeError:
            log_cb("❌ 错误：无法读取 ZOSAPI 枚举。")
            return False

        for i in range(1, num_surf - 1):
            try:
                surf = lde.GetSurfaceAt(i)
                mat_name = str(surf.Material).strip().upper()
                if mat_name == "" or mat_name == "MIRROR": continue

                solve_data = surf.MaterialCell.CreateSolveType(substitute_enum)
                try:
                    sub_solve = solve_data._S_MaterialSubstitute
                    sub_solve.Catalog = target_catalog
                except Exception as ex_prop:
                    log_cb(f"⚠️ Surface {i} 库设置警告: {ex_prop}")

                surf.MaterialCell.SetSolveData(solve_data)
                count += 1
                log_cb(f"  -> Surface {i}: 已设为 Substitute")
            except Exception as e:
                log_cb(f"⚠️ Surface {i} 设置失败: {e}")

        log_cb(f"✅ 操作完成：共设置 {count} 个表面为 Substitute 状态。")
        return True

    # =================================================================
    #  【强力覆盖版】 安全读取 + 强制设变量 (修复属性错误)
    # =================================================================
    def set_nonzero_aspheric_variables(self, log_cb):
        if not self.is_connected:
            log_cb("❌ 错误：系统未连接")
            return False

        lde = self.system.LDE
        num_surf = lde.NumberOfSurfaces
        count_surf = 0
        count_vars = 0
        
        log_cb("--- 正在扫描非球面 (强力覆盖模式) ---")

        try:
            SolveType = self.ZOSAPI.Editors.SolveType
            SurfaceColumn = self.ZOSAPI.Editors.LDE.SurfaceColumn
        except AttributeError as e:
            log_cb(f"❌ 错误：ZOSAPI 枚举加载失败: {e}")
            return False

        for i in range(1, num_surf): 
            try:
                surf = lde.GetSurfaceAt(i)
                s_type_str = str(surf.Type).strip()
                s_type_upper = s_type_str.upper()

                # 判定条件：ID "23" 或 名称含 "EVEN"
                is_target = (s_type_str == "23") or ("EVEN" in s_type_upper)

                if is_target:
                    has_change = False
                    # 遍历 Parameter 1 到 18
                    for p_idx in range(1, 19): 
                        col_name = f"Par{p_idx}"
                        if not hasattr(SurfaceColumn, col_name): continue

                        try:
                            col_enum = getattr(SurfaceColumn, col_name)
                            cell = surf.GetSurfaceCell(col_enum)
                            
                            try:
                                val = cell.DoubleValue
                            except Exception:
                                val = 0.0

                            if abs(val) > 1e-25:
                                try:
                                    solve_data = cell.CreateSolveType(SolveType.Variable)
                                    cell.SetSolveData(solve_data)
                                    count_vars += 1
                                    has_change = True
                                    log_cb(f"    ✅ Surface {i} | Par {p_idx} ({val:.2e}) -> 设为变量")
                                except Exception as set_err:
                                    log_cb(f"    ❌ 设置变量失败 Surf {i} Par {p_idx}: {set_err}")
                        except Exception:
                            pass
                    
                    if has_change: count_surf += 1

            except Exception as e:
                log_cb(f"⚠️ Surface {i} 外层错误: {e}")
        
        if count_surf == 0 and count_vars == 0:
            log_cb("⚠️ 未修改任何变量。请检查所有非球面系数是否都为 0。")
        else:
            log_cb(f"🎉 成功！在 {count_surf} 个表面设置了 {count_vars} 个变量。")
        return True

    def run_local_optimization(self, log_cb):
        if not self.is_connected:
            log_cb("❌ 错误：系统未连接")
            return False

        log_cb("💾 [安全模式] 正在保存系统...")
        try:
            self.system.Save()
        except:
            pass

        log_cb("🚀 [安全模式] 启动优化 (单核/单圈)...")
        
        opt_tool = None
        try:
            # 打开本地优化工具
            opt_tool = self.system.Tools.OpenLocalOptimization()
            if opt_tool is None:
                log_cb("❌ 无法打开优化工具 (Zemax可能正忙)")
                return False
            
            # --- 设置参数 ---
            try:
                Algo = self.ZOSAPI.Tools.Optimization.OptimizationAlgorithm
                Cycles = self.ZOSAPI.Tools.Optimization.OptimizationCycles
                
                # 使用阻尼最小二乘法
                opt_tool.Algorithm = Algo.DampedLeastSquares
                
                # 【关键修改 1】只跑 1 圈 (测试能否存活)
                opt_tool.Cycles = Cycles.Fixed_1_Cycle
                
                # 【关键修改 2】只用 1 个核心 (最稳定，虽然慢)
                opt_tool.NumberOfCores = 4
                
                log_cb("⚙️ 配置: 算法=DLS | 循环=1圈 | 核心=1 (单线程)")
            except Exception as setup_err:
                log_cb(f"⚠️ 设置失败: {setup_err}")
                if opt_tool: opt_tool.Close()
                return False
            
            # 记录初始值
            start_mf = opt_tool.InitialMeritFunction
            log_cb(f"📉 初始 MF: {start_mf:.6f}")
            log_cb("⏳ 计算中 (请勿操作 Zemax 界面)...")
            
            # 执行优化
            opt_tool.RunAndWaitForCompletion()
            
            # 获取结果 (使用修正后的属性名 CurrentMeritFunction)
            end_mf = opt_tool.CurrentMeritFunction
            
            # 关闭工具
            opt_tool.Close()
            
            log_cb(f"✅ 优化成功! MF: {start_mf:.6f} -> {end_mf:.6f}")
            return True

        except Exception as e:
            err_msg = str(e)
            if "Pipe ended" in err_msg or "管道已结束" in err_msg:
                log_cb("❌ 严重错误: Zemax 软件崩溃。")
                log_cb("   原因可能是: 1.光路结构极不稳定 2.变量导致光线跑飞")
            else:
                log_cb(f"❌ 运行错误: {err_msg}")
            
            # 尝试清理
            if opt_tool:
                try: opt_tool.Close()
                except: pass
            return False
    def calculate_mtf_30_avg(self, log_cb):
        if not self.is_connected: return None, "未连接"
        log_cb("正在执行 MTF 分析 (Freq=30, Wave=2)...")
        try:
            num_waves = self.system.SystemData.Wavelengths.NumberOfWavelengths
            target_wave_idx = 2 if num_waves >= 2 else 1
            analysis = self.system.Analyses.New_Analysis(self.ZOSAPI.Analysis.AnalysisIDM.FftMtf)
            settings = analysis.GetSettings()
            try:
                settings.MaximumFrequency = 30.0
                settings.Wavelength.SetWavelengthNumber(target_wave_idx)
            except Exception: pass
            analysis.ApplyAndWaitForCompletion()
            
            results = analysis.GetResults()
            temp_file = os.path.join(tempfile.gettempdir(), "zos_mtf_temp.txt")
            results.GetTextFile(temp_file)
            
            mtf_values = []
            if os.path.exists(temp_file):
                with open(temp_file, 'r', encoding='utf-16-le') as f:
                    lines = f.readlines()
                for line in lines:
                    parts = line.strip().split()
                    if len(parts) < 3: continue
                    try:
                        freq = float(parts[0])
                        if abs(freq - 30.0) < 0.01:
                            val_t = float(parts[1])
                            val_s = float(parts[2])
                            mtf_values.append((val_t + val_s) / 2.0)
                    except ValueError: continue
                analysis.Close()
                try: os.remove(temp_file)
                except: pass
                
                if mtf_values:
                    final_avg = sum(mtf_values) / len(mtf_values)
                    log_cb(f"✅ MTF 计算完成: {final_avg:.4f} (基于 {len(mtf_values)} 个视场)")
                    return final_avg, len(mtf_values)
            return None, 0
        except Exception as e:
            log_cb(f"❌ MTF 分析过程出错: {e}")
            return None, 0
    
    # 在 ZemaxDataEngine 类中添加
    def run_smart_diagnosis(self, worst_offenders_text, log_cb):
        if not self.is_connected:
            log_cb("❌ 错误：系统未连接")
            return "系统未连接"
            
        log_cb("🤖 启动智能标签识别引擎...")
        
        # 实例化分析器
        labeler = SmartLabeler(self.system, self.ZOSAPI, log_cb)
        
        # 运行分析
        try:
            report = labeler.analyze(worst_offenders_text)
            log_cb("✅ 诊断报告生成完毕。")
            return report
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            log_cb(f"❌ 分析过程出错: {str(e)}")
            return f"分析出错: \n{err}"
    def generate_full_report(self, progress_cb, log_cb):
        if not self.is_connected: return "❌ 系统未连接"
        report_lines = []
        log_cb("正在初始化数据读取...")
        progress_cb(5, "初始化...")
        
        try:
            sys_name = self.system.SystemData.Title or "Untitled"
            eff_focal = self.system.SystemData.EFLE
            f_num = self.system.SystemData.ImageSpaceFNumber
            log_cb(f"读取系统参数: EFL={eff_focal:.2f}, F/#={f_num:.2f}")
        except Exception as e:
            sys_name = "Unknown"; eff_focal = 0; f_num = 0

        report_lines.append("================================================================")
        report_lines.append(f" 📑 ZEMAX LENS REPORT - {time.strftime('%Y-%m-%d %H:%M')}")
        report_lines.append(f" File: {sys_name} | EFL: {eff_focal:.4f} | F/#: {f_num:.4f}")
        report_lines.append("================================================================\n")

        lde = self.system.LDE
        num_surf = lde.NumberOfSurfaces
        
        report_lines.append(" [1] 镜头结构数据 (Lens Data)")
        report_lines.append(f" {'Surf':<4} | {'Type':<10} | {'Radius':<12} | {'Thickness':<12} | {'Glass':<10} | {'Semi-Dia':<10}")
        report_lines.append("-" * 75)

        step_progress = 40 / max(num_surf, 1)

        for i in range(num_surf):
            if i % 3 == 0: log_cb(f"正在读取表面 {i+1}/{num_surf} (LDE)...")
            progress_cb(10 + (i * step_progress), f"读取 LDE 表面 {i}...")

            try:
                surf = lde.GetSurfaceAt(i)
                s_type = str(surf.Type)
                if "Standard" in s_type: s_type = "Std"
                elif "EvenAsph" in s_type: s_type = "Asphere"
                mat = str(surf.Material).strip().upper()
                radius = surf.Radius
                thick = surf.Thickness
                sd = surf.SemiDiameter
                r_str = "Infinity" if abs(radius) > 1e8 else f"{radius:.4f}"
                report_lines.append(f" {i:<4} | {s_type:<10} | {r_str:<12} | {thick:<12.4f} | {mat:<10} | {sd:<10.4f}")
            except Exception: pass

        report_lines.append("\n")
        log_cb("✅ LDE 数据读取完成")

        progress_cb(60, "启动赛德尔分析...")
        try:
            analysis = self.system.Analyses.New_Analysis(self.ZOSAPI.Analysis.AnalysisIDM.SeidelCoefficients)
            analysis.ApplyAndWaitForCompletion()
            results = analysis.GetResults()
            temp_file = os.path.join(tempfile.gettempdir(), "zos_seidel.txt")
            results.GetTextFile(temp_file)
            
            if os.path.exists(temp_file):
                with open(temp_file, 'r', encoding='utf-16-le') as f: content = f.read()
                report_lines.append(" [2] 赛德尔像差系数 (Seidel Coefficients)")
                report_lines.append("-" * 90)
                capture = False
                seidel_data = []
                for line in content.splitlines():
                    if "Surf" in line and "S1" in line:
                        header = f" {'Surf':<4} | {'S1 (球差)':<12} | {'S2 (彗差)':<12} | {'S3 (像散)':<12} | {'S4 (场曲)':<12} | {'S5 (畸变)':<12}"
                        seidel_data.append(header); seidel_data.append("-" * 90)
                        capture = True; continue
                    if capture:
                        if line.strip() == "": break 
                        if "Sum" in line:
                            seidel_data.append("-" * 90)
                            parts = line.split()
                            if len(parts) > 5:
                                seidel_data.append(f" {'SUM':<4} | {parts[1]:<12} | {parts[2]:<12} | {parts[3]:<12} | {parts[4]:<12} | {parts[5]:<12}")
                            break
                        parts = line.split()
                        if len(parts) >= 6 and parts[0].isdigit():
                            seidel_data.append(f" {parts[0]:<4} | {parts[1]:<12} | {parts[2]:<12} | {parts[3]:<12} | {parts[4]:<12} | {parts[5]:<12}")
                report_lines.extend(seidel_data)
                analysis.Close()
                try: os.remove(temp_file)
                except: pass
        except Exception: pass

        progress_cb(85, "正在计算 MTF...")
        report_lines.append("\n")
        report_lines.append(" [3] MTF 分析报告")
        report_lines.append(f" 参数: Frequency=30 lp/mm | Wavelength Index=2 | Average of All Fields")
        report_lines.append("-" * 90)
        
        mtf_val, field_count = self.calculate_mtf_30_avg(log_cb)
        if mtf_val is not None:
            report_lines.append(f" 🟢 平均 MTF 值: {mtf_val:.4f}")
            report_lines.append(f"    (数据来源: {field_count} 个视场在 30lp/mm 处的 T/S 均值)")
        else:
            report_lines.append(" 🔴 MTF 获取失败或无数据")

        progress_cb(100, "生成完成")
        return "\n".join(report_lines)
    
# =================================================================
#  第二阶段核心：绝热演化控制器 (Adiabatic Evolution Controller) - v8.0 终极破局完整版
# =================================================================
class AdiabaticEvolutionController:
    def __init__(self, engine, log_cb):
        self.engine = engine
        self.system = engine.system
        self.ZOSAPI = engine.ZOSAPI
        self.log_cb = log_cb
        self.mfe = self.system.MFE
        self.lde = self.system.LDE
        
        # 1. 贪婪自适应动力学参数
        self.dynamic_lambda = 0.05   # 初始起步降敏 5%
        self.min_lambda = 0.005      # 最小退让到 0.5% (遇到阻力时)
        self.max_lambda = 0.20       # 顺风局最大单次砍 20%
        self.feather_weight = 0.05   # 破冰期的柔性牵引力权重
        
        self.max_generations = 30
        self.mtf_guard_threshold = 0.10 # 允许 10% 的阵痛期波动
        
        # 2. 非对称打击目标 (只打源头，彻底破坏旧有的病态平衡)
        self.dominant_targets = [
            {"name": "Front_Source", "surf": 3, "types": ["DIST", "ASTI"]}, 
            {"name": "Rear_Source", "surf": 19, "types": ["ASTI"]}      
        ]
        
        self.evolution_tasks = [] 
        self.current_target_index = 0

    def _insert_op(self, row, op_type, i1=None, i2=None, v1=None, v2=None, v3=None, v4=None, tgt=0.0, wgt=0.0):
        """ 
        ⚙️ 智能参数写入引擎 (终极防崩溃版)
        严格检查输入参数，并在写入 ZOS-API 单元格时捕获 .NET 异常。
        绝对不会再因为遇到灰色禁用单元格而导致程序卡死。
        """
        import System 
        op = self.mfe.InsertNewOperandAt(row)
        op.ChangeType(op_type)
        
        # 只有在明确传入参数时才去修改该单元格
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
        """ [物理层] 从 Analysis 接口校准真实赛德尔系数 """
        real_values = {}
        try:
            analysis = self.system.Analyses.New_Analysis(self.ZOSAPI.Analysis.AnalysisIDM.SeidelCoefficients)
            analysis.ApplyAndWaitForCompletion()
            results = analysis.GetResults()
            import tempfile
            import os
            tmp = tempfile.mktemp(suffix=".txt")
            results.GetTextFile(tmp)
            
            if os.path.exists(tmp):
                with open(tmp, 'r', encoding='utf-16-le') as f:
                    lines = f.readlines()
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 6 and parts[0].isdigit():
                        idx = int(parts[0])
                        try:
                            # 提取绝对值用于单边猛砸
                            real_values[idx] = {
                                "SPHA": abs(float(parts[1])),
                                "COMA": abs(float(parts[2])),
                                "ASTI": abs(float(parts[3])),
                                "DIST": abs(float(parts[5]))
                            }
                        except: pass
                try: os.remove(tmp)
                except: pass
            analysis.Close()
        except Exception as e:
            self.log_cb(f"⚠️ 读取赛德尔系数失败: {e}")
            
        return real_values

    def pre_emptive_freedom(self):
        """ 🔓 战前准备：强制解锁物理空间杠杆 """
        self.log_cb("🔓 [战前准备] 强制解锁物理自由度 (赋予降敏空间)...")
        SolveType = self.ZOSAPI.Editors.SolveType
        count_thick = 0
        count_rad = 0
        
        # 1. 解锁敏感面 [后] 的空气层厚度 (即当前面的 Thickness)
        for idx in [3, 4, 19, 20]:
            try:
                surf = self.lde.GetSurfaceAt(idx)
                if surf and surf.ThicknessCell.SolveType != SolveType.Variable:
                    surf.ThicknessCell.SetSolveData(surf.ThicknessCell.CreateSolveType(SolveType.Variable))
                    count_thick += 1
            except: pass
            
        # 2. 彻底释放海绵面 (面10, 23) 的曲率和厚度，用于吸收转移的像差
        for idx in [10, 23]:
            try:
                surf = self.lde.GetSurfaceAt(idx)
                if surf and surf.RadiusCell.SolveType != SolveType.Variable:
                    surf.RadiusCell.SetSolveData(surf.RadiusCell.CreateSolveType(SolveType.Variable))
                    count_rad += 1
                if surf and surf.ThicknessCell.SolveType != SolveType.Variable:
                    surf.ThicknessCell.SetSolveData(surf.ThicknessCell.CreateSolveType(SolveType.Variable))
                    count_thick += 1
            except: pass
            
        self.log_cb(f"   -> 成功夺取 {count_thick} 个厚度变量，{count_rad} 个曲率变量！")

    def setup_icebreaker_tank(self):
        """ 🛡️ 战前准备：部署几何防线与单边打击阵列 """
        self.log_cb("🛡️ [防线部署] 构建靶向几何防线与非对称打击阵列...")
        self.evolution_tasks = []
        OpType = self.ZOSAPI.Editors.MFE.MeritOperandType
        current_row = 1 
        
        try:
            last_surf = self.lde.NumberOfSurfaces - 1
            # ---------------------------------------------------
            # 1. 基础物理底线 (逐面厚度保护，防相交防变薄)
            # ---------------------------------------------------
            for i in range(1, last_surf):
                surf = self.lde.GetSurfaceAt(i)
                mat = str(surf.Material).upper().strip()
                if mat and mat != "AIR": # 玻璃实体
                    self._insert_op(current_row, OpType.CTGT, i1=i, tgt=0.5, wgt=2.0)
                    current_row += 1
                    self._insert_op(current_row, OpType.ETGT, i1=i, tgt=0.5, wgt=2.0)
                    current_row += 1
                else: # 空气间隔
                    self._insert_op(current_row, OpType.CTGT, i1=i, tgt=0.1, wgt=1.0)
                    current_row += 1

            # ---------------------------------------------------
            # 2. 核心病因打击：全视场入射角 (AOI) 双重熔断
            # ---------------------------------------------------
            for surf in [3, 4, 19, 20]:
                if surf >= last_surf: continue
                # RAID (边缘光线入射角): Wave=0, Hy=1.0, Py=1.0
                row_raid_m = current_row
                self._insert_op(row_raid_m, OpType.RAID, i1=surf, i2=0, v2=1.0, v4=1.0, tgt=0.0, wgt=0.0)
                current_row += 1
                # OPLT: 强制边缘入射角 < 45度
                self._insert_op(current_row, OpType.OPLT, i1=row_raid_m, tgt=45.0, wgt=5.0) 
                current_row += 1
                
                # RAIP (主光线入射角): Wave=0, Hy=1.0, Py=0.0
                row_raid_c = current_row
                self._insert_op(row_raid_c, OpType.RAID, i1=surf, i2=0, v2=1.0, v4=0.0, tgt=0.0, wgt=0.0)
                current_row += 1
                # OPLT: 强制主光线入射角 < 45度
                self._insert_op(current_row, OpType.OPLT, i1=row_raid_c, tgt=45.0, wgt=5.0) 
                current_row += 1

            # ---------------------------------------------------
            # 3. 非对称靶向打击阵列 (直接取绝对值，放弃SUMM净值)
            # ---------------------------------------------------
            true_values = self.get_true_seidel_values()
            op_map = {"SPHA": OpType.SPHA, "COMA": OpType.COMA, "ASTI": OpType.ASTI, "DIST": OpType.DIST}
            
            for target in self.dominant_targets:
                surf = target["surf"]
                if surf not in true_values: continue
                
                task_indices = []
                for t in target["types"]:
                    if t not in op_map: continue
                    real_val = true_values[surf].get(t, 0.0)
                    
                    # 读取原始像差
                    row_aberr = current_row
                    self._insert_op(row_aberr, op_map[t], i1=surf, i2=0, tgt=0.0, wgt=0.0)
                    current_row += 1
                    
                    # 取绝对值
                    row_abso = current_row
                    self._insert_op(row_abso, OpType.ABSO, i1=row_aberr, tgt=0.0, wgt=0.0)
                    current_row += 1
                    
                    # 植入柔性控制 OPLT
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
            
            # 刷新初始目标真值
            for task in self.evolution_tasks:
                real_net_val = self.mfe.GetOperandAt(task["abso_row"]).Value
                task["current_target"] = real_net_val
                task["op_object"].Target = real_net_val
                self.log_cb(f"   -> 锁定源头: Surf {task['surf']} [{task['type']}] = {real_net_val:.4f}")
                
            self.log_cb("✅ 几何破冰阵列部署完成！")
            
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self.log_cb(f"❌ 构建模块发生异常: \n{err}")

    def run_evolution(self):
        """ 🚀 几何破冰与贪婪自适应演化流程 """
        self.pre_emptive_freedom() # 强制给足变量
        
        if not self.evolution_tasks:
            self.setup_icebreaker_tank()
            
        initial_mtf, _ = self.engine.calculate_mtf_30_avg(lambda x: None)
        self.log_cb(f"🏁 初始 MTF: {initial_mtf:.4f} | 初始步长: {self.dynamic_lambda:.1%}")
        current_mtf = initial_mtf
        
        for gen in range(1, self.max_generations + 1):
            target = self.dominant_targets[self.current_target_index]
            self.current_target_index = (self.current_target_index + 1) % len(self.dominant_targets)
            
            self.log_cb(f"\n🌊 [第 {gen} 代] 非对称打击: {target['name']} (当前步长: {self.dynamic_lambda:.1%})")
            
            # --- 步骤 1：施加不平衡高压 ---
            for t_idx in target.get("task_indices", []):
                task = self.evolution_tasks[t_idx]
                old_tgt = task["current_target"]
                
                # 记住上一次的安全点，用于失败时回滚
                task["safe_target"] = old_tgt 
                
                # 贪婪砍价
                new_tgt = old_tgt * (1.0 - self.dynamic_lambda)
                task["current_target"] = new_tgt
                
                # 更新操作数
                task["op_object"].Target = new_tgt
                task["op_object"].Weight = self.feather_weight # 极羽量级权重
                self.log_cb(f"   🎯 单压 {task['type']}@{task['surf']}: {old_tgt:.4f} -> {new_tgt:.4f}")

            # 暂时放过其他源头 (休眠不打的组)
            for i, task in enumerate(self.evolution_tasks):
                if i not in target.get("task_indices", []):
                    task["op_object"].Weight = 0.0

            # --- 步骤 2：底层放权 (闭门猛跑) ---
            self.run_quick_focus()
            
            # 让Zemax跑 15 圈，充分利用多核和刚才释放的变量，去消化不平衡的指令
            success = self.run_quick_optimization(cycles=15) 
            self.run_quick_focus()
            
            if not success:
                self.log_cb("❌ [致命熔断] 物理边界被突破，光线无法追迹或系统畸变过大！")
                self.log_cb("⛔ 任务强制安全终止。系统已锁定在崩溃前最后一次有效状态。")
                break 
                
            # --- 步骤 3：贪婪自适应判定 ---
            new_mtf, _ = self.engine.calculate_mtf_30_avg(lambda x: None)
            drop_rate = (current_mtf - new_mtf) / current_mtf if current_mtf > 0 else 0
            
            if drop_rate < 0.02: 
                # MTF 坚挺 (下降极小甚至提升) -> 系统游刃有余
                self.log_cb(f"   📈 MTF 坚挺 ({new_mtf:.4f})。系统游刃有余！")
                # 贪婪加速：放大步长 1.5 倍
                self.dynamic_lambda = min(self.dynamic_lambda * 1.5, self.max_lambda)
                current_mtf = new_mtf
                
            elif drop_rate > self.mtf_guard_threshold: 
                # MTF 暴跌 -> 遭遇硬骨头
                self.log_cb(f"   ⚠️ MTF 暴跌至 {new_mtf:.4f} (Drop: {drop_rate:.1%})。触发回滚与深吸气！")
                
                # 1. 目标回滚 (Undo 施压，承认刚才步子迈大了)
                for t_idx in target.get("task_indices", []):
                    tk_obj = self.evolution_tasks[t_idx]
                    tk_obj["current_target"] = tk_obj["safe_target"]
                    tk_obj["op_object"].Target = tk_obj["safe_target"]
                    
                # 2. 步长缩减 (认怂)
                self.dynamic_lambda = max(self.dynamic_lambda * 0.5, self.min_lambda)
                
                # 3. 呼吸式恢复 (让海绵变量发挥作用)
                self.respiratory_recovery() 
                
                recov_mtf, _ = self.engine.calculate_mtf_30_avg(lambda x: None)
                self.log_cb(f"   ✨ 海绵区吸水完毕，恢复至: {recov_mtf:.4f}")
                
                # 无论如何更新基准
                current_mtf = max(recov_mtf, new_mtf) 
                
            else:
                # 正常波动
                self.log_cb(f"   📊 稳步推进，MTF: {new_mtf:.4f} (Drop: {drop_rate:.1%})")
                current_mtf = new_mtf
            
        if success:
            self.log_cb("\n✅ 破局降敏圆满结束！请检查海绵面 (10, 23) 和空气间隔的吸收情况。")

    def run_quick_focus(self):
        """ 执行快速散焦补偿 """
        try:
            qf = self.system.Tools.OpenQuickFocus()
            if qf:
                qf.RunAndWaitForCompletion()
                qf.Close()
        except: pass

    def run_quick_optimization(self, cycles=15):
        """ 闭门优化引擎：包含光线熔断检查 """
        try:
            tool = self.system.Tools.OpenLocalOptimization()
            if tool:
                tool.Algorithm = self.ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares
                tool.Cycles = self.ZOSAPI.Tools.Optimization.OptimizationCycles.Fixed_10_Cycles # 扩大底层计算量
                tool.NumberOfCores = 8 # 拉满核数，加快运算
                tool.RunAndWaitForCompletion()
                final_mf = tool.CurrentMeritFunction
                tool.Close()
                
                # 【崩溃判定条件】如果大于 1e8 或 NaN，判定为光线全反射或溢出
                if final_mf > 1e8 or final_mf < 0: 
                    return False
                return True
        except: 
            return False
        return False

    def respiratory_recovery(self):
        """ 呼吸式像质回补：全面放松像差约束，全功率恢复 MTF """
        # 将所有像差约束权重降至极低，甚至接近 0
        for task in self.evolution_tasks:
            task["op_object"].Weight = self.feather_weight * 0.1 
            
        self.log_cb("   💨 [深呼吸] 卸载源头压力，让海绵区全力工作恢复像质...")
        self.run_quick_focus()
        # 闭门狂奔 20 圈，纯靠默认评价函数里的 MTF/Spot 控制把画质拉回来
        self.run_quick_optimization(cycles=20)
        
# =================================================================
#  前端 UI：Seidel Reporter (带 Log)
# =================================================================
class SeidelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zemax Seidel & Lens Reporter Pro")
        self.root.geometry("1000x850")
        self.root.configure(bg="#F3F4F6")
        
        self.engine = ZemaxDataEngine()
        self.is_running = False

        self._init_styles()
        self._build_ui()

    def _init_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TProgressbar", thickness=10, background="#2563EB", borderwidth=0)
        
        self.colors = {
            "primary": "#2563EB", 
            "dark": "#1F2937", 
            "bg": "#F3F4F6",
            "success": "#059669",
            "purple": "#7C3AED", 
            "orange": "#EA580C",
            "danger": "#DC2626", # 新增红色用于优化
            "terminal_bg": "#111827", 
            "terminal_fg": "#4ADE80"  
        }
        self.fonts = {
            "head": ("Impact", 18),
            "mono": ("Consolas", 10),
            "btn": ("微软雅黑", 10, "bold"),
            "log": ("Consolas", 9)
        }

    def _build_ui(self):
        # --- 1. 侧边栏 ---
        sidebar = tk.Frame(self.root, bg=self.colors["dark"], width=200)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="ZEMAX\nREPORTER", font=self.fonts["head"], fg="white", bg=self.colors["dark"]).pack(pady=40)

        self.status_indicator = tk.Canvas(sidebar, width=30, height=30, bg=self.colors["dark"], highlightthickness=0)
        self.status_indicator.pack(pady=(0, 5))
        self.light = self.status_indicator.create_oval(5, 5, 25, 25, fill="#4B5563", outline="")
        
        self.btn_connect = tk.Button(sidebar, text="🔗 连接 OpticStudio", bg="#374151", fg="white", 
                                     relief="flat", font=("微软雅黑", 9), command=self.connect_zemax)
        self.btn_connect.pack(fill=tk.X, padx=20, pady=10)

        tk.Frame(sidebar, height=1, bg="#4B5563").pack(fill=tk.X, padx=10, pady=20)

        # --- 绝热演化按钮 ---
        tk.Frame(sidebar, height=1, bg="#4B5563").pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_evolution = tk.Button(sidebar, text="🧬 启动绝热演化", bg="#7C3AED", fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.action_run_evolution)
        self.btn_evolution.pack(fill=tk.X, padx=20, pady=5)
        
        # --- 功能按钮 ---
        self.btn_run = tk.Button(sidebar, text="🚀 生成分析报告", bg=self.colors["primary"], fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.run_report)
        self.btn_run.pack(fill=tk.X, padx=20)

        self.btn_set_s = tk.Button(sidebar, text="🔧 一键设为替换 (S)", bg=self.colors["purple"], fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.action_set_substitute)
        self.btn_set_s.pack(fill=tk.X, padx=20, pady=10)

        self.btn_set_asph = tk.Button(sidebar, text="💎 设非球面变量", bg=self.colors["orange"], fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.action_set_asph_vars)
        self.btn_set_asph.pack(fill=tk.X, padx=20, pady=10)
        
        self.btn_optimize = tk.Button(sidebar, text="⚡ 执行优化", bg=self.colors["danger"], fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.action_run_optimization)
        self.btn_optimize.pack(fill=tk.X, padx=20, pady=10)

        # --- 智能诊断按钮 ---
        tk.Frame(sidebar, height=1, bg="#4B5563").pack(fill=tk.X, padx=10, pady=10) # 分隔线
        
        self.btn_smart = tk.Button(sidebar, text="🤖 智能标签诊断", bg="#0891B2", fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.action_smart_diagnosis)
        self.btn_smart.pack(fill=tk.X, padx=20, pady=5)

        self.btn_save = tk.Button(sidebar, text="💾 导出 TXT", bg="#10B981", fg="white", 
                                 relief="flat", font=self.fonts["btn"], pady=10, state=tk.DISABLED,
                                 command=self.save_to_file)
        self.btn_save.pack(fill=tk.X, padx=20, pady=20)

        # --- 2. 主区域 ---
        main = tk.Frame(self.root, bg=self.colors["bg"])
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=20)

        header = tk.Frame(main, bg="white", pady=15, padx=20)
        header.pack(fill=tk.X)
        tk.Label(header, text="镜头数据与像差分析 (Lens & Aberration Report)", font=("微软雅黑", 12, "bold"), fg="#374151", bg="white").pack(anchor="w")

        self.progress = ttk.Progressbar(main, style="TProgressbar", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(15, 5))
        self.lbl_status = tk.Label(main, text="等待连接...", bg=self.colors["bg"], fg="#6B7280", font=("微软雅黑", 9))
        self.lbl_status.pack(anchor="w")

        # 使用 PanedWindow 管理区域
        paned_window = tk.PanedWindow(main, orient=tk.VERTICAL, bg=self.colors["bg"], sashwidth=6, sashrelief="ridge")
        paned_window.pack(fill=tk.BOTH, expand=True, pady=10)

        # [区域 1] 报告显示区
        report_frame = tk.Frame(paned_window, bg="white")
        report_scroll = tk.Scrollbar(report_frame)
        report_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_output = tk.Text(report_frame, font=self.fonts["mono"], yscrollcommand=report_scroll.set, 
                                  bg="white", fg="#1F2937", padx=15, pady=15, relief="flat", height=15)
        self.txt_output.pack(fill=tk.BOTH, expand=True)
        report_scroll.config(command=self.txt_output.yview)
        paned_window.add(report_frame, minsize=200)

        # [区域 2] 文件加载与预览区 (修改部分)
        input_frame = tk.Frame(paned_window, bg="white")
        
        # 顶部工具栏
        tool_bar = tk.Frame(input_frame, bg="#E5E7EB", height=30, padx=5, pady=5)
        tool_bar.pack(fill=tk.X)
        
        tk.Label(tool_bar, text="公差数据源:", bg="#E5E7EB", fg="#374151", font=("微软雅黑", 9, "bold")).pack(side=tk.LEFT, padx=(5,10))
        
        # 加载按钮
        self.btn_load_tol = tk.Button(tool_bar, text="📂 加载报告文件 (.txt)", bg="white", fg="#2563EB", 
                                      font=("微软雅黑", 8), relief="groove", padx=10, 
                                      command=self.action_load_tolerance_file)
        self.btn_load_tol.pack(side=tk.LEFT)
        
        tk.Label(tool_bar, text="(支持 UTF-16/UTF-8, 自动提取 Worst Offenders)", bg="#E5E7EB", fg="#6B7280", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=10)

        # 预览文本框
        self.txt_input_tol = tk.Text(input_frame, height=6, font=("Consolas", 9), relief="flat", bg="#F9FAFB", fg="#374151", padx=10, pady=5)
        self.txt_input_tol.pack(fill=tk.BOTH, expand=True)
        self.txt_input_tol.insert(tk.END, ">>> 请点击上方按钮加载 Zemax 公差报告文件，或在此处直接粘贴数据...")
        
        paned_window.add(input_frame, minsize=100)

        # [区域 3] 日志区
        log_frame = tk.Frame(paned_window, bg=self.colors["terminal_bg"])
        log_header = tk.Frame(log_frame, bg="#374151", height=25, padx=10)
        log_header.pack(fill=tk.X)
        tk.Label(log_header, text="💻 调试终端 (Debug Log)", fg="#D1D5DB", bg="#374151", font=("Consolas", 8)).pack(side=tk.LEFT)
        log_scroll = tk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log = tk.Text(log_frame, font=self.fonts["log"], yscrollcommand=log_scroll.set,
                               bg=self.colors["terminal_bg"], fg=self.colors["terminal_fg"], 
                               padx=10, pady=10, relief="flat", height=6, state=tk.DISABLED)
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.txt_log.yview)
        paned_window.add(log_frame, minsize=80)

    def connect_zemax(self):
        self.log_to_terminal("正在尝试连接 Zemax API...")
        self.lbl_status.config(text="正在连接 Zemax API...", fg=self.colors["primary"])
        self.root.update()
        success, msg = self.engine.connect()
        if success:
            self.status_indicator.itemconfig(self.light, fill="#10B981") 
            self.btn_connect.config(text="已连接", state=tk.DISABLED, bg="#065F46")
            self._set_buttons_state(tk.NORMAL)
            self.lbl_status.config(text=f"就绪: {msg}", fg="#059669")
            self.log_to_terminal(f"连接成功: {msg}")
        else:
            self.status_indicator.itemconfig(self.light, fill="#EF4444") 
            messagebox.showerror("连接失败", msg)
            self.lbl_status.config(text="连接失败", fg="#DC2626")
            self.log_to_terminal(f"连接失败: {msg}")
    
    # -----------------------------------------------------------
    #  【新增】智能诊断按钮触发逻辑
    # -----------------------------------------------------------
    def action_smart_diagnosis(self):
        if self.is_running: return
        
        # 获取用户在中间文本框粘贴的内容
        tol_text = self.txt_input_tol.get(1.0, tk.END).strip()
        if not tol_text or "TCON" not in tol_text and "TTHI" not in tol_text:
            messagebox.showwarning("提示", "请先在中间的输入框粘贴 Worst Offenders 数据！\n(需要包含 TCON, TTHI 等关键字)")
            return
    

    # =================================================================
    #  【新增】 按钮响应：加载公差文件
    # =================================================================
    def action_load_tolerance_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Zemax 公差报告文件",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return

        # 调用后端读取
        extracted_text = self.engine.load_tolerance_file(file_path, self.log_to_terminal)
        
        if extracted_text:
            # 清空并填入预览框
            self.txt_input_tol.delete(1.0, tk.END)
            self.txt_input_tol.insert(tk.END, extracted_text)
            self.log_to_terminal(f"文件加载成功，准备就绪。")
            # 自动激活智能诊断按钮
            self.btn_smart.config(state=tk.NORMAL)
        else:
            messagebox.showwarning("解析失败", "无法从文件中提取 'Worst Offenders' 数据。\n请确保文件格式正确。")

    # =================================================================
    #  【智能诊断】 按钮响应逻辑
    # =================================================================
    def action_smart_diagnosis(self):
        if self.is_running: return
        
        # 1. 获取用户在中间文本框粘贴(或加载)的内容
        tol_text = self.txt_input_tol.get(1.0, tk.END).strip()
        
        # 简单检查内容是否有效
        if not tol_text or ("TCON" not in tol_text and "TTHI" not in tol_text):
            messagebox.showwarning("提示", "请先加载或粘贴 Worst Offenders 数据！\n(需要包含 TCON, TTHI 等关键字)")
            return

        # 2. 定义后台工作线程函数 (注意缩进：它在 action_smart_diagnosis 内部)
        def _smart_worker():
            # 调用引擎的诊断功能
            report = self.engine.run_smart_diagnosis(tol_text, self.log_to_terminal)
            
            # 定义 UI 更新函数 (注意缩进：它在 _smart_worker 内部)
            def _update_ui():
                # 将结果显示在上方的大报告框中
                self.txt_output.delete(1.0, tk.END)
                self.txt_output.insert(tk.END, "====== 🧬 智能光学标签与处方报告 🧬 ======\n\n")
                self.txt_output.insert(tk.END, report)
                
                self.lbl_status.config(text="诊断完成", fg=self.colors["success"])
                self._set_buttons_state(tk.NORMAL)
                
            # 在主线程更新 UI
            self.root.after(0, _update_ui)

        # 3. 启动线程 (注意缩进：它属于 action_smart_diagnosis，与 _smart_worker 定义平级)
        self._set_buttons_state(tk.DISABLED)
        self.log_to_terminal("--- 开始智能诊断 ---")
        threading.Thread(target=_smart_worker, daemon=True).start()

    def action_run_evolution(self):
        if self.is_running: return
        
        # 简单的参数确认弹窗
        if not messagebox.askyesno("确认", "即将开始【绝热演化】流程。\n\n这将大幅修改评价函数 (MFE) 并进行多轮优化。\n建议先备份文件。\n\n是否继续？"):
            return

        def _evo_worker():
            # 实例化控制器
            controller = AdiabaticEvolutionController(self.engine, self.log_to_terminal)
            
            # 运行演化
            controller.run_evolution()
            
            # 完成后刷新界面
            self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("完成", "绝热演化流程已结束。"))

        self._set_buttons_state(tk.DISABLED)
        self.log_to_terminal("--- 启动绝热演化 (Adiabatic Evolution) ---")
        threading.Thread(target=_evo_worker, daemon=True).start()

    def _set_buttons_state(self, state):
        # 基础功能按钮
        self.btn_run.config(state=state)
        self.btn_set_s.config(state=state)
        self.btn_set_asph.config(state=state)
        self.btn_optimize.config(state=state)
        
        # 智能诊断相关
        self.btn_smart.config(state=state)
        
        # 【关键修复】确保绝热演化按钮也被启用
        self.btn_evolution.config(state=state) 
        
        # 导出按钮逻辑保持不变
        if state == tk.NORMAL and self.txt_output.get(1.0, tk.END).strip():
            self.btn_save.config(state=tk.NORMAL)
        else:
            self.btn_save.config(state=tk.DISABLED)

    def run_report(self):
        if self.is_running: return
        self.is_running = True
        self._set_buttons_state(tk.DISABLED)
        self.txt_output.delete(1.0, tk.END)
        self._clear_log()
        self.log_to_terminal("任务开始：准备生成报告...")
        threading.Thread(target=self._worker, daemon=True).start()

    def action_set_substitute(self):
        if self.is_running: return
        def _set_s_worker():
            self.engine.set_all_materials_substitute(self.log_to_terminal)
            self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("完成", "已将玻璃设置为 Substitute (S) 模式"))
        self._set_buttons_state(tk.DISABLED)
        self.log_to_terminal("--- 开始批量修改玻璃属性 ---")
        threading.Thread(target=_set_s_worker, daemon=True).start()

    def action_set_asph_vars(self):
        if self.is_running: return
        def _set_asph_worker():
            self.engine.set_nonzero_aspheric_variables(self.log_to_terminal)
            self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("完成", "非零非球面系数已设为变量"))
        self._set_buttons_state(tk.DISABLED)
        self.log_to_terminal("--- 开始设置非球面变量 ---")
        threading.Thread(target=_set_asph_worker, daemon=True).start()

    # =================================================================
    #  【新增】 触发优化操作
    # =================================================================
    def action_run_optimization(self):
        if self.is_running: return
        
        def _opt_worker():
            self.engine.run_local_optimization(self.log_to_terminal)
            self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("完成", "优化已完成"))

        self._set_buttons_state(tk.DISABLED)
        self.log_to_terminal("--- 开始执行本地优化 ---")
        threading.Thread(target=_opt_worker, daemon=True).start()

    def log_to_terminal(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] > {msg}\n"
        def _write():
            self.txt_log.config(state=tk.NORMAL)
            self.txt_log.insert(tk.END, formatted_msg)
            self.txt_log.see(tk.END) 
            self.txt_log.config(state=tk.DISABLED)
        self.root.after(0, _write)

    def _clear_log(self):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def _worker(self):
        def update_prog(val, msg):
            self.root.after(0, lambda: self._update_ui_progress(val, msg))
        report = self.engine.generate_full_report(update_prog, self.log_to_terminal)
        self.root.after(0, lambda: self._finish_report(report))

    def _update_ui_progress(self, val, msg):
        self.progress['value'] = val
        self.lbl_status.config(text=msg)

    def _finish_report(self, report):
        self.txt_output.insert(tk.END, report)
        self.is_running = False
        self._set_buttons_state(tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)
        self.lbl_status.config(text="报告生成完毕", fg=self.colors["success"])
        self.log_to_terminal("任务结束。")

    def save_to_file(self):
        content = self.txt_output.get(1.0, tk.END)
        if not content.strip(): return
        f = filedialog.asksaveasfilename(defaultextension=".txt", 
                                         filetypes=[("Text Files", "*.txt")],
                                         initialfile="Lens_Report.txt")
        if f:
            with open(f, 'w', encoding='utf-8') as file:
                file.write(content)
            self.log_to_terminal(f"文件已保存至: {f}")
            messagebox.showinfo("导出成功", f"文件已保存:\n{f}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SeidelApp(root)
    root.mainloop()