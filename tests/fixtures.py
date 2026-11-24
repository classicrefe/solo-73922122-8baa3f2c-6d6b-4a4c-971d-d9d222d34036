# -*- coding: utf-8 -*-
"""测试夹具：构造符合/违反业务规则的送检 XLSX。"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


HEADERS = ["序号", "采样站码", "批号", "唯一样本号", "检测项目", "结果", "备注公式列"]


def build_demo_workbook(path: str) -> None:
    """构造标准演示表：

    批次 B2401（站 S01，3 行）、B2402（站 S02，2 行）、B2403（站 S01，2 行）。
    B2401 与 B2403 站码相同但跨批，下载时不得跨批合并。
    站码列纵向合并；结果列含公式，验证公式原样保留。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "送检表"

    ws.append(HEADERS)
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="D9E1F2")
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = bold
        cell.fill = fill

    # (批号, 样本号, 站码, 结果值)
    data = [
        ("B2401", "SMP-001", "S01", 10),
        ("B2401", "SMP-002", "S01", 20),
        ("B2401", "SMP-003", "S01", 30),
        ("B2402", "SMP-004", "S02", 40),
        ("B2402", "SMP-005", "S02", 50),
        ("B2403", "SMP-006", "S01", 60),
        ("B2403", "SMP-007", "S01", 70),
    ]
    for idx, (batch, sample, station, value) in enumerate(data, start=1):
        r = idx + 1
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=station)
        ws.cell(row=r, column=3, value=batch)
        ws.cell(row=r, column=4, value=sample)
        ws.cell(row=r, column=5, value="pH")
        ws.cell(row=r, column=6, value=value)
        # 公式列：引用结果列，验证导出后公式原样保留
        ws.cell(row=r, column=7, value=f"=F{r}*2")

    # 站码列纵向合并：2-4、5-6、7-8
    ws.merge_cells("B2:B4")
    ws.merge_cells("B5:B6")
    ws.merge_cells("B7:B8")
    ws.cell(row=2, column=2).alignment = Alignment(vertical="center")

    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["D"].width = 14
    wb.save(path)


def build_duplicate_workbook(path: str) -> None:
    """样本号重复表：SMP-002 在第 3、5 行重复。"""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    data = [
        ("B2401", "SMP-001", "S01"),
        ("B2401", "SMP-002", "S01"),
        ("B2401", "SMP-003", "S01"),
        ("B2402", "SMP-002", "S02"),  # 重复
        ("B2402", "SMP-005", "S02"),
    ]
    for idx, (batch, sample, station) in enumerate(data, start=1):
        r = idx + 1
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=station)
        ws.cell(row=r, column=3, value=batch)
        ws.cell(row=r, column=4, value=sample)
    ws.merge_cells("B2:B4")
    ws.merge_cells("B5:B6")
    wb.save(path)


def build_split_batch_workbook(path: str) -> None:
    """同批不连续：B2401 被 B2402 打断。"""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    data = [
        ("B2401", "SMP-001", "S01"),
        ("B2401", "SMP-002", "S01"),
        ("B2402", "SMP-003", "S02"),
        ("B2401", "SMP-004", "S01"),
    ]
    for idx, (batch, sample, station) in enumerate(data, start=1):
        r = idx + 1
        ws.cell(row=r, column=1, value=idx)
        ws.cell(row=r, column=2, value=station)
        ws.cell(row=r, column=3, value=batch)
        ws.cell(row=r, column=4, value=sample)
    wb.save(path)


def build_missing_header_workbook(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    # 三列关键字全部缺失
    ws.append(["序号", "监测点位名称", "批次", "样品编码"])
    ws.append([1, "S01", "B1", "X1"])
    wb.save(path)
