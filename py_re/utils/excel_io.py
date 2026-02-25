import time
import pandas as pd
from typing import Optional, List, Union, Tuple
import xlwings as xw
from config.constants import EXCEL_FONT, EXCEL_FONT_SIZE, PIVOT_FREEZE_CELL, NORMAL_FREEZE_CELL
from utils.logger import write_log, log_error
from data_model.asmp import AsmpRow
from data_model.loss import LossRow

class ExcelIO:
    """Excel IO工具类（批量读写+静默重试+模型转换）"""
    RETRY_COUNT: int = 3
    RETRY_DELAY: float = 0.5

    @classmethod
    def read_data(
        cls, wb: xw.Book, ws_name: str, header: bool = False
    ) -> Optional[pd.DataFrame]:
        """
        批量读取Excel数据为DataFrame（支持COM重试）
        :param wb: 工作簿对象
        :param ws_name: 工作表名
        :param header: 是否包含表头
        :return: DataFrame（读取失败返回None）
        """
        write_log("ExcelIO", "read_data", f"开始读取工作表: {ws_name}", "Excel IO")
        for retry in range(cls.RETRY_COUNT):
            try:
                time.sleep(0.1)  # 给COM接口响应时间
                sht = wb.sheets[ws_name]
                # 批量读取（比逐行读取快10倍+）
                df = sht.range("A1").expand().options(
                    pd.DataFrame, index=False, header=header
                ).value
                write_log("ExcelIO", "read_data", f"读取成功: {ws_name} ({len(df)}行)", "Excel IO")
                return df
            except Exception as e:
                # 工作表不存在直接抛出
                if ws_name not in [s.name for s in wb.sheets]:
                    err_msg = f"工作表不存在: {ws_name}"
                    log_error("ExcelIO", "read_data", err_msg)
                    raise ValueError(err_msg)
                # 重试机制
                if retry == cls.RETRY_COUNT - 1:
                    err_msg = f"读取工作表失败（重试{cls.RETRY_COUNT}次）: {str(e)}"
                    log_error("ExcelIO", "read_data", err_msg)
                    raise
                time.sleep(cls.RETRY_DELAY)
        return None

    @classmethod
    def read_asmp_data(cls, wb: xw.Book, ws_name: str) -> List[AsmpRow]:
        """读取假设数据并转换为AsmpRow模型列表"""
        df = cls.read_data(wb, ws_name, header=False)
        if df is None or df.empty:
            return []
        # 跳过表头（假设数据第1行为表头）
        return [AsmpRow.from_list(row) for row in df.iloc[1:].values.tolist()]

    @classmethod
    def read_loss_data(cls, wb: xw.Book, ws_name: str) -> List[LossRow]:
        """读取损失数据并转换为LossRow模型列表"""
        df = cls.read_data(wb, ws_name, header=False)
        if df is None or df.empty:
            return []
        return [LossRow.from_list(row) for row in df.values.tolist()]

    @classmethod
    def write_data(
        cls, wb: xw.Book, ws_name: str, data: Union[pd.DataFrame, List[list]], is_pivot: bool = False
    ) -> bool:
        """
        批量写入数据到Excel（覆盖模式）
        :param wb: 工作簿对象
        :param ws_name: 目标工作表名
        :param data: 要写入的数据（DataFrame或list）
        :param is_pivot: 是否为透视表
        :return: 写入成功与否
        """
        write_log("ExcelIO", "write_data", f"开始写入工作表: {ws_name}", "Excel IO")
        try:
            # 转换为DataFrame（统一写入逻辑）
            if isinstance(data, list):
                df = pd.DataFrame(data[1:], columns=data[0])
            else:
                df = data.copy()

            # 获取/创建工作表
            if ws_name in [s.name for s in wb.sheets]:
                sht = wb.sheets[ws_name]
                sht.clear()
            else:
                sht = wb.sheets.add(ws_name, after=wb.sheets.count)

            # 批量写入（比逐行写入快5倍+）
            sht.range("A1").value = df

            # 格式化工作表（对齐VBA FormatSheetStandard）
            cls._format_sheet(sht, is_pivot)

            write_log("ExcelIO", "write_data", f"写入成功: {ws_name} ({len(df)}行)", "Excel IO")
            return True
        except Exception as e:
            log_error("ExcelIO", "write_data", f"写入失败: {str(e)}")
            return False

    @classmethod
    def _format_sheet(cls, sht: xw.Sheet, is_pivot: bool) -> None:
        """格式化工作表（对齐VBA标准）"""
        used_range = sht.range("A1").expand()
        # 字体配置
        used_range.api.Font.Name = EXCEL_FONT
        used_range.api.Font.Size = EXCEL_FONT_SIZE

        # 表头样式
        header_range = sht.range(1, 1).resize(1, used_range.columns.count)
        header_range.api.HorizontalAlignment = -4108  # 居中
        header_range.api.Font.Bold = True
        header_range.api.Interior.ColorIndex = 15  # 灰色背景

        # 冻结窗格
        sht.activate()
        sht.api.Application.ActiveWindow.FreezePanes = False
        sht.range(PIVOT_FREEZE_CELL if is_pivot else NORMAL_FREEZE_CELL).select()
        sht.api.Application.ActiveWindow.FreezePanes = True

        # 列宽自适应
        if is_pivot:
            sht.range("A:J").autofit()
            sht.range("J:J").number_format = "#,##0.00"
        else:
            # 数值列格式化
            headers = sht.range(1, 1).expand("right").value
            for col_idx, header in enumerate(headers, 1):
                header_str = str(header).lower() if header else ""
                if any(keyword in header_str for keyword in ["amount", "loss", "recovery", "retention", "limit"]):
                    sht.range(2, col_idx).expand("down").number_format = "#,##0.00"
                elif "share" in header_str:
                    sht.range(2, col_idx).expand("down").number_format = "0%"
            sht.autofit()

    @classmethod
    def read_named_range_value(cls, wb: xw.Book, named_range: str) -> Optional[Union[str, int, float]]:
        """读取VBA命名区域的值（带异常处理）"""
        try:
            return wb.names[named_range].refers_to_range.value
        except Exception as e:
            log_error("ExcelIO", "read_named_range_value", f"读取命名区域失败: {named_range}, {str(e)}")
            return None