from typing import Final

# 日志配置
LOG_FILE_NAME: Final[str] = "calculation_engine.log"
LOG_FORMAT: Final[str] = "[%(asctime)s] [%(module)s.%(funcName)s] %(levelname)s - %(message)s"
LOG_ENCODING: Final[str] = "utf-8-sig"

# 计算默认值
DEFAULT_MAX_YEARS: Final[int] = 10000
DEFAULT_ACTIVE_LIMIT: Final[float] = 1e15
DEFAULT_MATERIALITY_THRESHOLD: Final[float] = 0.5

# Excel配置
EXCEL_FONT: Final[str] = "Arial"
EXCEL_FONT_SIZE: Final[int] = 10
PIVOT_FREEZE_CELL: Final[str] = "K2"
NORMAL_FREEZE_CELL: Final[str] = "A2"

# 临时文件配置
TEMP_CSV_PREFIX: Final[str] = "_temp.csv"