# data_model/asmp.py
from dataclasses import dataclass
from typing import Optional, List
from config.schema import SCHEMA

@dataclass(frozen=True)
class AsmpRow:
    """假设数据行模型（对应ASMP SCHEMA）"""
    tid: str
    segment: str
    sub_lob: str
    peril: str
    limit_risk: Optional[float]
    limit_event: Optional[float]
    retention: Optional[float]
    share: Optional[float]
    inuring: Optional[int]
    limit_form: str
    group: Optional[str]

    @classmethod
    def from_list(cls, row: list) -> "AsmpRow":
        """从VBA传入的list转换为模型对象"""
        schema = SCHEMA["ASMP"]
        return cls(
            tid=str(row[schema["tid"]]).strip() if row[schema["tid"]] else "",
            segment=str(row[schema["segment"]]).strip() if row[schema["segment"]] else "",
            sub_lob=str(row[schema["sub_lob"]]).strip() if row[schema["sub_lob"]] else "",
            peril=str(row[schema["peril"]]).strip() if row[schema["peril"]] else "",
            limit_risk=float(row[schema["limit_risk"]]) if row[schema["limit_risk"]] else 0.0,
            limit_event=float(row[schema["limit_event"]]) if row[schema["limit_event"]] else 0.0,
            retention=float(row[schema["retention"]]) if row[schema["retention"]] else 0.0,
            share=float(row[schema["share"]]) if row[schema["share"]] else 0.0,
            inuring=int(row[schema["inuring"]]) if row[schema["inuring"]] else 0,
            limit_form=str(row[schema["limit_form"]]).strip() if row[schema["limit_form"]] else "Risk",
            group=str(row[schema["group"]]).strip() if row[schema["group"]] else ""
        )

@dataclass
class AsmpRule:
    """计算规则模型（简化核心字段）"""
    tid: str
    active_limit: float
    retention: float
    share: float
    inuring: int
    group: str
    limit_risk: float
    limit_event: float