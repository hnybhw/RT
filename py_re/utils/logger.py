import logging
import os
from typing import Optional, Union
import xlwings as xw
from config.constants import LOG_FILE_NAME, LOG_FORMAT, LOG_ENCODING
from config.schema import NAMED_RANGES

class ExcelLogger:
    """日志工具类（同时输出到Debug窗口+Excel Setup表+文件）"""
    def __init__(self, wb: xw.Book):
        self.wb = wb
        self.log_enabled = self._get_log_switch()
        self.logger = self._setup_logger()
        self.log_start_cell = self._get_log_start_cell()

    def _get_log_switch(self) -> bool:
        """读取VBA日志开关（Main!p_export_xllog）"""
        try:
            log_val = self.wb.names[NAMED_RANGES["LOG_CONTROL"]].refers_to_range.value
            return str(log_val).strip().upper() in ["YES", "TRUE", "1"]
        except Exception:
            return False

    def _get_log_start_cell(self) -> xw.Range:
        """获取Excel日志起始单元格（Setup!E2）"""
        try:
            setup_sht = self.wb.sheets["2@Setup"]  # VBA定义的永久表名
            return setup_sht.range("E2")
        except Exception:
            raise ValueError("未找到Setup工作表或日志起始单元格E2")

    def _setup_logger(self) -> logging.Logger:
        """配置日志（文件+控制台/Debug窗口）"""
        logger = logging.getLogger("RI_Engine")
        logger.setLevel(logging.INFO)
        logger.propagate = False

        # 清理现有handler
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        # 控制台handler（输出到VBA立即窗口）
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter(LOG_FORMAT))
        logger.addHandler(console_handler)

        # 文件handler（仅日志开关开启时）
        if self.log_enabled:
            try:
                log_path = os.path.join(os.path.dirname(self.wb.fullname), LOG_FILE_NAME)
                file_handler = logging.FileHandler(
                    log_path, mode="a", encoding=LOG_ENCODING
                )
                file_handler.setFormatter(logging.Formatter(LOG_FORMAT))
                logger.addHandler(file_handler)
            except Exception as e:
                logger.error(f"日志文件创建失败: {str(e)}")

        return logger

    def _get_last_log_row(self) -> int:
        """获取Excel日志最后一行"""
        last_row = self.log_start_cell.end("down").row
        return last_row if last_row > self.log_start_cell.row else self.log_start_cell.row

    def write_log(
        self, module_name: str, proc_name: str, log_content: str, log_title: str = ""
    ) -> None:
        """
        对齐VBA WriteLog方法的日志写入
        :param module_name: 模块名
        :param proc_name: 过程/函数名
        :param log_content: 日志内容
        :param log_title: 日志分类标题（写入F列）
        """
        # 格式化日志内容（与VBA日志格式一致）
        log_text = f"[{module_name}.{proc_name}] {log_content}"
        
        # 输出到Debug窗口+文件
        self.logger.info(log_text)

        # 写入Excel（日志开关开启时）
        if self.log_enabled:
            try:
                last_row = self._get_last_log_row() + 1
                self.log_start_cell.sheet.range(f"E{last_row}").value = log_text
                self.log_start_cell.sheet.range(f"F{last_row}").value = log_title
            except Exception as e:
                self.logger.error(f"Excel日志写入失败: {str(e)}")

    def error(self, module_name: str, proc_name: str, error_msg: str, log_title: str = "错误信息") -> None:
        """错误日志（带堆栈信息）"""
        self.logger.error(f"[{module_name}.{proc_name}] {error_msg}", exc_info=True)
        if self.log_enabled:
            try:
                last_row = self._get_last_log_row() + 1
                self.log_start_cell.sheet.range(f"E{last_row}").value = f"[ERROR] [{module_name}.{proc_name}] {error_msg}"
                self.log_start_cell.sheet.range(f"F{last_row}").value = log_title
            except Exception:
                pass

# 全局日志实例（通过init_logger初始化）
global_logger: Optional[ExcelLogger] = None

def init_logger(wb: xw.Book) -> None:
    """初始化全局日志（在main函数中调用）"""
    global global_logger
    global_logger = ExcelLogger(wb)

def write_log(module_name: str, proc_name: str, log_content: str, log_title: str = "") -> None:
    """全局日志写入函数（简化调用）"""
    if global_logger:
        global_logger.write_log(module_name, proc_name, log_content, log_title)

def log_error(module_name: str, proc_name: str, error_msg: str) -> None:
    """全局错误日志写入"""
    if global_logger:
        global_logger.error(module_name, proc_name, error_msg)