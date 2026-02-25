import xlwings as xw
import traceback
from utils.logger import init_logger, write_log, log_error
from orchestration.task_orchestrator import TaskOrchestrator

def main() -> None:
    """统一主入口（VBA通过xlwings调用）"""
    wb = xw.Book.caller()
    try:
        # 初始化日志（对齐VBA日志规则）
        init_logger(wb)
        write_log("RI_Engine", "main", "Python再保险引擎启动", "系统启动")

        # 创建编排器并执行
        orchestrator = TaskOrchestrator(wb)
        orchestrator.run()

        write_log("RI_Engine", "main", "Python再保险引擎执行完成", "系统结束")

    except Exception as e:
        error_detail = traceback.format_exc()
        log_error("RI_Engine", "main", f"引擎崩溃: {str(e)}\n{error_detail}")
        # 弹出错误提示（兼容VBA交互）
        wb.app.alert(f"Python 再保险引擎执行失败，详细错误:\n\n{error_detail}")
    finally:
        wb.app.status_bar = False

if __name__ == "__main__":
    main()