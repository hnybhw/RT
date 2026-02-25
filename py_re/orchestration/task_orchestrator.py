import time
import gc
from typing import Dict, List, Optional
import xlwings as xw
from utils.excel_io import ExcelIO
from utils.logger import write_log, log_error
from config.schema import NAMED_RANGES
from config.constants import DEFAULT_MAX_YEARS
from data_process.asmp_processor import AsmpProcessor
from calculation.engine_factory import CalculationEngineFactory
from pivot.pivot_converter import PivotConverter

class TaskOrchestrator:
    """任务编排器（原orchestrator逻辑，分层解耦）"""
    def __init__(self, wb: xw.Book):
        self.wb = wb
        self.excel_io = ExcelIO()
        self.asmp_processor = AsmpProcessor()
        self.pivot_converter = PivotConverter(wb)
        self.data_pool: Dict[str, List[list]] = {}  # 数据池（跨任务共享数据）

    def load_config(self) -> Tuple[Optional[str], Optional[List[list]], int]:
        """加载VBA配置（命名区域读取）"""
        # 1. 读取目标分段
        target_segment = self.excel_io.read_named_range_value(
            self.wb, NAMED_RANGES["TARGET_SEGMENT"]
        )
        if not target_segment:
            log_error("TaskOrchestrator", "load_config", "未读取到目标分段（ref_10K_Segment）")
            return None, None, DEFAULT_MAX_YEARS

        # 2. 读取任务配置
        cfg_range = self.excel_io.read_named_range_value(
            self.wb, NAMED_RANGES["CONFIG_RANGE"]
        )
        if not cfg_range or not isinstance(cfg_range, list):
            log_error("TaskOrchestrator", "load_config", "未读取到任务配置（rng_config）")
            return target_segment, None, DEFAULT_MAX_YEARS

        # 3. 读取max_year（从VBA主控表p_sim_yrs读取）
        max_years_raw = self.excel_io.read_named_range_value(
            self.wb, NAMED_RANGES["MAX_YEARS"]
        )
        max_years = int(max_years_raw) if max_years_raw and str(max_years_raw).isdigit() else DEFAULT_MAX_YEARS

        write_log("TaskOrchestrator", "load_config", f"配置加载完成: 分段={target_segment}, 最大年份={max_years}", "配置管理")
        return target_segment, cfg_range, max_years

    def execute_task(self, task: Dict[str, str], max_years: int) -> None:
        """执行单个任务"""
        task_name = task.get("Task_Name", "未知任务")
        func_name = task.get("Function_Name")
        input_var = task.get("Input_Var")
        asmp_sheet = task.get("Asmp_Sheet")
        params = task.get("Params", "")
        output_var = task.get("Output_Var")
        save_sheet = task.get("Save_Sheet")
        print_flag = task.get("Print", "NO").upper() == "YES"

        write_log("TaskOrchestrator", "execute_task", f"开始执行任务: {task_name}", "任务执行")
        self.wb.app.status_bar = f"执行中: {task_name}"

        try:
            # 分支执行不同任务
            if func_name == "fetch_data_loss":
                # 读取损失数据
                loss_data = self.excel_io.read_data(self.wb, params)
                result = loss_data.values.tolist() if loss_data is not None else []

            elif func_name == "fetch_data_asmp":
                # 读取假设数据
                asmp_rows = self.excel_io.read_asmp_data(self.wb, params)
                result = [row.to_list() for row in asmp_rows]

            elif func_name == "calculate_reinsurance_engine":
                # 核心计算（工厂模式创建引擎）
                p_list = [p.strip() for p in params.split(",") if p.strip()]
                contract_type = p_list[1] if len(p_list) > 1 else "XL"
                engine = CalculationEngineFactory.create_engine(contract_type)

                # 获取输入数据
                loss_list = self.data_pool.get(input_var, [])
                asmp_list = self.data_pool.get(asmp_sheet, [])

                # 转换为模型对象
                loss_rows = [self.excel_io.loss_row_from_list(row) for row in loss_list[1:]]  # 跳过表头
                asmp_rows = [self.excel_io.asmp_row_from_list(row) for row in asmp_list]
                asmp_rules = self.asmp_processor.build_asmp_rules(asmp_rows, p_list[0])

                # 执行计算
                result = engine.calculate(loss_rows, asmp_rules, p_list[0])

            elif func_name == "pivot_recovery_to_wide":
                # 透视表转换
                recovery_data = self.data_pool.get(input_var, [])
                result = self.pivot_converter.convert_to_wide(recovery_data, save_sheet, max_years)

            elif func_name in ["aggregate_events", "split_recovery_by_audit_trail", "get_net_by_seq", "apply_qs_event_cap", "apply_max_contribution_cap"]:
                # 其他计算任务（调用对应模块）
                result = self._execute_other_tasks(func_name, input_var, asmp_sheet, params)

            else:
                raise ValueError(f"未定义的任务函数: {func_name}")

            # 存入数据池
            if output_var:
                self.data_pool[output_var] = result
                write_log("TaskOrchestrator", "execute_task", f"任务结果存入数据池: {output_var}", "数据管理")

            # 写入Excel（按需）
            if print_flag and save_sheet and result:
                self.excel_io.write_data(self.wb, save_sheet, result, is_pivot="pivot" in func_name)

        except Exception as e:
            log_error("TaskOrchestrator", "execute_task", f"任务执行失败: {task_name}, 错误: {str(e)}")
            raise
        finally:
            self.wb.app.status_bar = False

    def _execute_other_tasks(self, func_name: str, input_var: str, asmp_sheet: str, params: str) -> List[list]:
        """执行其他辅助计算任务"""
        from calculation.aggregator import EventAggregator
        from calculation.recovery_allocator import RecoveryAllocator
        from calculation.cap_processor import QSEventCapProcessor, MaxContributionCapProcessor

        input_data = self.data_pool.get(input_var, [])
        asmp_data = self.data_pool.get(asmp_sheet, []) if asmp_sheet else []

        if func_name == "aggregate_events":
            p_list = [p.strip() for p in params.split(",")]
            aggregator = EventAggregator()
            return aggregator.aggregate(input_data, asmp_data, p_list[0], p_list[1])

        elif func_name == "split_recovery_by_audit_trail":
            ref_data = self.data_pool.get(params.strip(), [])
            allocator = RecoveryAllocator()
            return allocator.split_by_audit_trail(ref_data, input_data)

        elif func_name == "get_net_by_seq":
            p_list = [p.strip() for p in params.split(",")]
            cap_processor = QSEventCapProcessor()
            return cap_processor.get_net_by_seq(input_data, self.data_pool.get(p_list[0], []), int(p_list[1]))

        elif func_name == "apply_qs_event_cap":
            cap_processor = QSEventCapProcessor()
            return cap_processor.apply_qs_event_cap(input_data, asmp_data)

        elif func_name == "apply_max_contribution_cap":
            cap_processor = MaxContributionCapProcessor()
            return cap_processor.apply_max_cap(input_data, float(params))

        else:
            raise ValueError(f"未实现的辅助任务: {func_name}")

    def run(self) -> None:
        """核心运行入口"""
        # 禁用Excel屏幕刷新和自动计算（提升性能）
        app = self.wb.app
        app.screen_updating = False
        app.calculation = "manual"

        try:
            # 加载配置
            target_segment, cfg_range, max_years = self.load_config()
            if not target_segment or not cfg_range:
                return

            # 解析任务配置
            tasks = [dict(zip(cfg_range[0], row)) for row in cfg_range[1:] if row[0] and str(row[1]).strip().upper() == target_segment.upper()]
            if not tasks:
                write_log("TaskOrchestrator", "run", f"无匹配任务: {target_segment}", "任务执行")
                return

            # 执行所有任务
            for task in tasks:
                self.execute_task(task, max_years)

            # 清理数据池（保留核心数据）
            self._clean_data_pool()

            write_log("TaskOrchestrator", "run", f"所有任务执行完成: {target_segment}", "任务执行")

        except Exception as e:
            log_error("TaskOrchestrator", "run", f"编排执行失败: {str(e)}")
            raise
        finally:
            # 恢复Excel环境
            app.screen_updating = True
            app.calculation = "automatic"
            app.status_bar = False

    def _clean_data_pool(self) -> None:
        """清理数据池（释放内存）"""
        # 仅保留最后一个任务的输出数据
        if self.data_pool:
            last_output_var = list(self.data_pool.keys())[-1]
            self.data_pool = {last_output_var: self.data_pool[last_output_var]}
        gc.collect()
        write_log("TaskOrchestrator", "_clean_data_pool", "数据池清理完成", "内存管理")