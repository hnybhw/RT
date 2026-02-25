# data_model/loss.py
from dataclasses import dataclass
from typing import Optional
from config.schema import SCHEMA

@dataclass(frozen=True)
class LossRow:
    """损失数据行模型（对应LOSS SCHEMA）"""
    seq: Optional[int]
    colume: Optional[int]
    year: Optional[int]
    sub_lob: str
    peril: str
    event: str
    amount: float
    tid: str
    group: str

    @classmethod
    def from_list(cls, row: list) -> "LossRow":
        schema = SCHEMA["LOSS"]
        return cls(
            seq=int(float(row[schema["seq"]])) if row[schema["seq"]] else None,
            colume=int(row[schema["colume"]]) if row[schema["colume"]] else None,
            year=int(row[schema["year"]]) if row[schema["year"]] else None,
            sub_lob=str(row[schema["sub_lob"]]).strip() if row[schema["sub_lob"]] else "",
            peril=str(row[schema["peril"]]).strip() if row[schema["peril"]] else "",
            event=str(row[schema["event"]]).strip() if row[schema["event"]] else "",
            amount=float(row[schema["amount"]]) if row[schema["amount"]] else 0.0,
            tid=str(row[schema["tid"]]).strip() if row[schema["tid"]] else "",
            group=str(row[schema["group"]]).strip() if row[schema["group"]] else ""
        )