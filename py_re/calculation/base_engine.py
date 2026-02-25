from abc import ABC, abstractmethod
from typing import List, Tuple
from data_model.loss import LossRow
from data_model.asmp import AsmpRule
from data_model.recovery import RecoveryRow

class BaseCalculationEngine(ABC):
    """计算引擎抽象基类（策略模式）"""
    @abstractmethod
    def calculate(
        self, loss_rows: List[LossRow], asmp_rules: List[AsmpRule], loss_segment: str
    ) -> List[RecoveryRow]:
        """
        核心计算接口
        :param loss_rows: 损失数据列表
        :param asmp_rules: 假设规则列表
        :param loss_segment: 业务分段
        :return: 赔付结果列表
        """
        pass

    @staticmethod
    def _get_active_limit(limit_form: str, limit_risk: float, limit_event: float) -> float:
        """获取有效限额（统一逻辑抽离）"""
        from config.constants import DEFAULT_ACTIVE_LIMIT
        if limit_form == "Risk":
            return limit_risk
        elif limit_form == "Event":
            return limit_event
        else:
            return DEFAULT_ACTIVE_LIMIT