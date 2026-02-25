from typing import Dict, Final

# 字段映射SCHEMA（严格对齐VBA数据列索引，避免魔法数字）
SCHEMA: Final[Dict[str, Dict[str, int]]] = {
    "ASMP": {
        "tid": 1, "segment": 3, "sub_lob": 5, "peril": 6,
        "limit_risk": 8, "limit_event": 9, "retention": 10,
        "share": 11, "inuring": 12, "limit_form": 13, "group": 14
    },
    "LOSS": {
        "seq": 0, "colume": 1, "year": 2, "sub_lob": 3,
        "peril": 4, "event": 5, "amount": 6, "tid": 7, "group": 8
    },
    "RECOVERY": {
        "tid": 0, "segment": 1, "seq": 2, "colume": 3, "year": 4,
        "sub_lob": 5, "peril": 6, "event": 7, "amount": 8,
        "limit_risk": 9, "limit_event": 10, "retention": 11,
        "share": 12, "inuring": 13, "recovery": 14, "group": 15
    }
}

# 命名区域配置（与VBA完全一致，便于维护）
NAMED_RANGES: Final[Dict[str, str]] = {
    "TARGET_SEGMENT": "ref_10K_Segment",
    "CONFIG_RANGE": "rng_config",
    "LOG_CONTROL": "p_export_xllog",  # 对齐VBA日志开关（原p_export_pylog废弃）
    "MAX_YEARS": "p_sim_yrs",         # 从VBA主控表读取max_year
    "MATERIALITY_THRESHOLD": "p_materiality_threshold"
}