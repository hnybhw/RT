from typing import Dict, Type
from calculation.base_engine import BaseCalculationEngine
from calculation.excess_of_loss import ExcessOfLossEngine
# 后续可扩展其他引擎（如ProportionalEngine）
from utils.logger import write_log

class CalculationEngineFactory:
    """计算引擎工厂（动态创建引擎实例）"""
    _engine_registry: Dict[str, Type[BaseCalculationEngine]] = {
        "XL": ExcessOfLossEngine,  # 超赔合同
        # "PROP": ProportionalEngine,  # 比例合同（后续扩展）
        # "SSC": SlidingScaleEngine   # 滑动佣金合同（后续扩展）
    }

    @classmethod
    def register_engine(cls, contract_type: str, engine_class: Type[BaseCalculationEngine]) -> None:
        """注册新引擎（扩展时使用）"""
        if contract_type not in cls._engine_registry:
            cls._engine_registry[contract_type] = engine_class
            write_log("CalculationEngineFactory", "register_engine", f"注册新引擎: {contract_type}", "工厂模式")

    @classmethod
    def create_engine(cls, contract_type: str = "XL") -> BaseCalculationEngine:
        """
        创建引擎实例
        :param contract_type: 合同类型（默认超赔XL）
        :return: 计算引擎实例
        """
        engine_class = cls._engine_registry.get(contract_type.upper())
        if not engine_class:
            raise ValueError(f"不支持的合同类型: {contract_type}，支持类型: {list(cls._engine_registry.keys())}")
        
        engine = engine_class()
        write_log("CalculationEngineFactory", "create_engine", f"创建引擎实例: {contract_type}", "工厂模式")
        return engine