import pandas as pd
import numpy as np
from typing import List, Dict, Tuple
from data_model.loss import LossRow
from data_model.asmp import AsmpRule
from data_model.recovery import RecoveryRow
from calculation.base_engine import BaseCalculationEngine
from config.schema import SCHEMA

class ExcessOfLossEngine(BaseCalculationEngine):
    """超赔合同计算引擎（原核心逻辑，Pandas向量化优化）"""
    def calculate(
        self, loss_rows: List[LossRow], asmp_rules: List[AsmpRule], loss_segment: str
    ) -> List[RecoveryRow]:
        """
        向量化计算超赔赔付（无显式for循环）
        :return: RecoveryRow列表（兼容原list输出格式）
        """
        # 转换为DataFrame（向量化基础）
        loss_df = self._loss_rows_to_df(loss_rows)
        rule_df = self._asmp_rules_to_df(asmp_rules)

        if loss_df.empty or rule_df.empty:
            return [RecoveryRow.get_header_row()]

        # 1. 规则匹配：先按tid匹配，再按(sub_lob, peril)兜底
        merged_df = self._match_rules(loss_df, rule_df)

        # 2. 向量化计算赔付金额
        merged_df["recovery"] = self._calculate_recovery(merged_df)

        # 3. 过滤有效结果（recovery>0）
        result_df = merged_df[merged_df["recovery"] > 0].copy()

        # 4. 转换为RecoveryRow列表（兼容原输出格式）
        return self._df_to_recovery_rows(result_df, loss_segment)

    def _loss_rows_to_df(self, loss_rows: List[LossRow]) -> pd.DataFrame:
        """LossRow列表转换为DataFrame"""
        data = [
            {
                "seq": row.seq,
                "colume": row.colume,
                "year": row.year,
                "sub_lob": row.sub_lob,
                "peril": row.peril,
                "event": row.event,
                "amount": row.amount,
                "tid": row.tid,
                "group": row.group
            }
            for row in loss_rows if row.seq is not None
        ]
        df = pd.DataFrame(data)
        # 类型优化（提升计算效率）
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0.0)
        df["tid_clean"] = df["tid"].str.strip()
        df["sub_lob_clean"] = df["sub_lob"].str.strip()
        df["peril_clean"] = df["peril"].str.strip()
        return df

    def _asmp_rules_to_df(self, asmp_rules: List[AsmpRule]) -> pd.DataFrame:
        """AsmpRule列表转换为DataFrame"""
        data = [
            {
                "tid": rule.tid,
                "limit_risk": rule.limit_risk,
                "limit_event": rule.limit_event,
                "active_limit": rule.active_limit,
                "retention": rule.retention,
                "share": rule.share,
                "inuring": rule.inuring,
                "group": rule.group,
                "tid_clean": rule.tid.strip(),
                "sub_lob_clean": rule.tid.split("_")[0] if "_" in rule.tid else "",  # 适配sub_lob匹配
                "peril_clean": rule.tid.split("_")[1] if "_" in rule.tid else ""
            }
            for rule in asmp_rules
        ]
        return pd.DataFrame(data)

    def _match_rules(self, loss_df: pd.DataFrame, rule_df: pd.DataFrame) -> pd.DataFrame:
        """向量化规则匹配（替代原dict查找）"""
        # 按tid匹配
        tid_merged = loss_df.merge(
            rule_df,
            left_on="tid_clean",
            right_on="tid_clean",
            how="left",
            suffixes=("", "_rule")
        )

        # 按(sub_lob, peril)兜底匹配
        sub_peril_merged = loss_df.merge(
            rule_df,
            left_on=["sub_lob_clean", "peril_clean"],
            right_on=["sub_lob_clean", "peril_clean"],
            how="left",
            suffixes=("", "_fallback")
        )

        # 合并匹配结果（优先tid匹配）
        merged_df = tid_merged.combine_first(sub_peril_merged)
        return merged_df.dropna(subset=["active_limit"])  # 过滤无匹配规则的记录

    def _calculate_recovery(self, merged_df: pd.DataFrame) -> pd.Series:
        """向量化计算赔付金额"""
        # 公式：min(max(损失金额-自留额, 0), 有效限额) * 分保比例
        return (
            np.minimum(
                np.maximum(merged_df["amount"] - merged_df["retention"], 0),
                merged_df["active_limit"]
            ) * merged_df["share"]
        ).round(4)

    def _df_to_recovery_rows(self, df: pd.DataFrame, loss_segment: str) -> List[RecoveryRow]:
        """DataFrame转换为RecoveryRow列表（兼容原输出格式）"""
        recovery_schema = SCHEMA["RECOVERY"]
        header = RecoveryRow.get_header_row()
        rows = [header]

        for _, row in df.iterrows():
            recovery_row = RecoveryRow(
                tid=row["tid"],
                segment=loss_segment,
                seq=row["seq"],
                colume=row["colume"],
                year=row["year"],
                sub_lob=row["sub_lob"],
                peril=row["peril"],
                event=row["event"],
                amount=row["amount"],
                limit_risk=row["limit_risk"],
                limit_event=row["limit_event"],
                retention=row["retention"],
                share=row["share"],
                inuring=row["inuring"],
                recovery=row["recovery"],
                group=row["group"]
            )
            rows.append(recovery_row.to_list())

        return rows