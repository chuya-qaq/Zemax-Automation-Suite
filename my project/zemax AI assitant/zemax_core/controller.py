# zemax_core/controller.py
import clr
import os
import sys

# 导入配置和各个服务模块
from .config import ZEMAX_PATH, Q_GRADES
from .services_tol import setup_q_grade_tolerances_impl, lock_all_semi_diameters_impl
from .services_ana import run_sensitivity_analysis_impl
from .services_ai import run_ai_heuristic_optimization_impl


# 确保能找到 Zemax 的 DLL
if os.path.exists(ZEMAX_PATH):
    sys.path.append(ZEMAX_PATH)
    os.environ['PATH'] = ZEMAX_PATH + os.pathsep + os.environ['PATH']

class ZemaxController:
    def __init__(self):
        self.app = None
        self.system = None
        self.ZOSAPI = None
        self.is_connected = False
        self.last_sensitivity_data = [] 
        self.Q_GRADES = Q_GRADES # 挂载配置供外部读取

    def _log(self, cb, msg):
        """辅助日志函数"""
        if cb: cb(msg)

    def connect(self):
        """
        连接逻辑保留在 Controller 内部，因为它是状态管理的核心。
        """
        try:
            if not os.path.exists(os.path.join(ZEMAX_PATH, "ZOSAPI_NetHelper.dll")):
                return False, f"找不到 DLL，请检查路径: {ZEMAX_PATH}"

            clr.AddReference(os.path.join(ZEMAX_PATH, "ZOSAPI_NetHelper.dll"))
            import ZOSAPI_NetHelper
            ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(ZEMAX_PATH)
            
            clr.AddReference(os.path.join(ZEMAX_PATH, "ZOSAPI.dll"))
            import ZOSAPI
            self.ZOSAPI = ZOSAPI 

            conn = ZOSAPI.ZOSAPI_Connection()
            self.app = conn.ConnectAsExtension(0)
            
            if self.app and self.app.PrimarySystem:
                self.system = self.app.PrimarySystem
                self.is_connected = True
                return True, f"已连接 | 镜头表面数: {self.system.LDE.NumberOfSurfaces}"
            
            return False, "连接失败：请确保 OpticStudio 已打开。"
        except Exception as e:
            return False, f"连接异常: {str(e)}"
   
    # =========================================================================
    #  核心逻辑委托区 (Delegation Area)
    #  这里所有的函数都只是“传声筒”，直接调用 services 里的最新实现。
    # =========================================================================

    def setup_q_grade_tolerances(self, grade="Q4", custom_frn=0.0, custom_thi=0.0, log_cb=None):
        """
        [公差建模] -> 委托给 services_tol
        """
        return setup_q_grade_tolerances_impl(self, grade, custom_frn, custom_thi, log_cb)

    def lock_all_semi_diameters(self, log_cb=None):
        """
        [锁定口径] -> 委托给 services_tol (含智能跳过逻辑)
        """
        return lock_all_semi_diameters_impl(self, log_cb)

    def run_sensitivity_analysis(self, progress_cb, log_cb=None, grade_settings=None):
        """
        [敏感度分析] -> 委托给 services_ana (含防崩溃 Safe Close 逻辑)
        """
        return run_sensitivity_analysis_impl(self, progress_cb, log_cb, grade_settings)

    def run_ai_heuristic_optimization(self, strategy_id, progress_cb, log_cb=None, constraints=None):
        """
        [AI 优化] -> 委托给 services_ai
        增加了 constraints 参数用于自定义 EFFL/WFNO
        """
        return run_ai_heuristic_optimization_impl(self, strategy_id, progress_cb, log_cb, constraints)