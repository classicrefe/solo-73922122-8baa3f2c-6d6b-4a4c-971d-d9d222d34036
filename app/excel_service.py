# -*- coding: utf-8 -*-
"""Excel 处理：站码合并格展开、校验、更正应用、结果合并与工作簿导出。

业务约定
--------
1. 送检表把同一批次、连续若干行的采样站码纵向合并在一个单元格中。
2. 批号、唯一样本号逐行填写（不合并）。
3. 同批行连续且站码相同；跨批即使站码相同也是不同段。

模块只依赖 openpyxl，不依赖 Flask，便于独立测试。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 表头候选关键字（按优先级），全部做去空白包含匹配
STATION_KEYS = ("采样站码", "站码", "站点编码", "站点代码", "采样站点")
BATCH_KEYS = ("批号", "批次号", "批次编号", "批编号")
SAMPLE_KEYS = ("唯一样本号", "样本号", "样品编号", "样本编号", "样品号", "样品唯一编号")


class WorkbookError(Exception):
    """工作簿处理/校验失败。"""


@dataclass
class ColumnSpec:
    """表头定位结果。"""

    header_row: int
    station_col: int
    batch_col: int
    sample_col: int
    headers: List[str]


@dataclass
class RowRecord:
    """一行数据的解析结果（行号为 1 基 openpyxl 行号）。"""

    row_index: int
    batch: str
    sample_no: str
    original_station: str
    effective_station: str = ""
    source: str = ""  # original / batch / sample / revoked->batch
    sample_correction: Optional[dict] = None
    batch_correction: Optional[dict] = None


@dataclass
class ParsedWorkbook:
    """解析后的工作簿上下文。"""

    path: str
    sheet_name: str
    spec: ColumnSpec
    rows: List[RowRecord] = field(default_factory=list)
    # 站码列原有的合并区间，key=(min_row, max_row)
    station_merges: List[Tuple[int, int]] = field(default_factory=list)
    # openpyxl 工作簿对象（仅上传流程内存中使用，不持久化）
    workbook: Any = None


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def _norm(value) -> str:
    """单元格值标准化：去首尾空白；None -> 空串。"""
    if value is None:
        return ""
    return str(value).strip()


def find_header(ws, keys: Tuple[str, ...]) -> Optional[Tuple[int, int, str]]:
    """在前 20 行内查找包含任一关键字的表头单元格。

    返回 (行号, 列号, 表头文本)，找不到返回 None。
    """
    max_scan = min(ws.max_row, 20)
    for row in range(1, max_scan + 1):
        for col in range(1, ws.max_column + 1):
            text = _norm(ws.cell(row=row, column=col).value)
            if not text:
                continue
            for key in keys:
                if key in text:
                    return row, col, text
    return None


def detect_columns(ws) -> ColumnSpec:
    """自动识别站码、批号、样本号所在列与表头行。"""
    station_hit = find_header(ws, STATION_KEYS)
    batch_hit = find_header(ws, BATCH_KEYS)
    sample_hit = find_header(ws, SAMPLE_KEYS)

    missing = []
    if station_hit is None:
        missing.append("采样站码")
    if batch_hit is None:
        missing.append("批号")
    if sample_hit is None:
        missing.append("唯一样本号")
    if missing:
        raise WorkbookError(
            "表头识别失败，缺少列：" + "、".join(missing)
            + "。请确认表头包含“采样站码/批号/唯一样本号”等关键字。"
        )

    header_row = max(station_hit[0], batch_hit[0], sample_hit[0])
    # 三个表头理应在同一行；给出版本化提示，避免错位
    rows = {station_hit[0], batch_hit[0], sample_hit[0]}
    if len(rows) != 1:
        raise WorkbookError(
            "采样站码、批号、唯一样本号表头不在同一行，无法可靠解析（检测到表头行："
            f"站码第{station_hit[0]}行、批号第{batch_hit[0]}行、"
            f"样本号第{sample_hit[0]}行）。"
        )

    headers = [
        _norm(ws.cell(row=header_row, column=c).value)
        for c in range(1, ws.max_column + 1)
    ]
    return ColumnSpec(
        header_row=header_row,
        station_col=station_hit[1],
        batch_col=batch_hit[1],
        sample_col=sample_hit[1],
        headers=headers,
    )


def _merged_ranges_for_col(ws, col: int) -> List[Tuple[int, int, int, int]]:
    """返回覆盖指定列的所有合并单元格边界 (min_row, min_col, max_row, max_col)。"""
    result = []
    letter = get_column_letter(col)
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        if min_col <= col <= max_col:
            result.append((min_row, min_col, max_row, max_col))
    return result


def unmerge_station_cells(ws, spec: ColumnSpec) -> List[Tuple[int, int]]:
    """展开站码列中数据区的纵向合并格，并把左上角值填入每一行。

    返回被展开的 (起始行, 结束行) 列表。其他列的合并格保持不动。
    站码列若存在横向合并（跨多列）也按业务异常拒收。
    """
    data_first = spec.header_row + 1
    expanded: List[Tuple[int, int]] = []
    for min_row, min_col, max_row, max_col in _merged_ranges_for_col(
        ws, spec.station_col
    ):
        if max_row < data_first:
            continue  # 表头或标题区的合并不动
        if min_col != max_col:
            raise WorkbookError(
                f"站码列第 {min_row}-{max_row} 行为横向合并单元格，"
                "送检表站码只允许纵向合并，请整理后重新上传。"
            )
        top_value = ws.cell(row=min_row, column=spec.station_col).value
        ws.unmerge_cells(
            start_row=min_row,
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )
        for r in range(min_row, max_row + 1):
            ws.cell(row=r, column=spec.station_col).value = top_value
        expanded.append((min_row, max_row))
    return expanded


def parse_rows(ws, spec: ColumnSpec) -> List[RowRecord]:
    """逐行读取数据区，做必填与样本号重复校验。"""
    data_first = spec.header_row + 1
    rows: List[RowRecord] = []
    seen: Dict[str, int] = {}
    duplicates: List[Tuple[str, int, int]] = []  # (样本号, 首次行, 重复行)

    for r in range(data_first, ws.max_row + 1):
        # 整行空白则视为表尾，跳过
        if all(
            _norm(ws.cell(row=r, column=c).value) == ""
            for c in range(1, ws.max_column + 1)
        ):
            continue
        batch = _norm(ws.cell(row=r, column=spec.batch_col).value)
        sample_no = _norm(ws.cell(row=r, column=spec.sample_col).value)
        station = _norm(ws.cell(row=r, column=spec.station_col).value)

        label = f"第 {r} 行"
        if not batch:
            raise WorkbookError(f"{label}缺少批号，批号必须逐行填写。")
        if not sample_no:
            raise WorkbookError(f"{label}缺少唯一样本号，样本号必须逐行填写。")
        if not station:
            raise WorkbookError(
                f"{label}缺少采样站码（含合并格展开后仍为空），请核对原表。"
            )
        if sample_no in seen:
            duplicates.append((sample_no, seen[sample_no], r))
        else:
            seen[sample_no] = r
        rows.append(
            RowRecord(
                row_index=r,
                batch=batch,
                sample_no=sample_no,
                original_station=station,
            )
        )

    if duplicates:
        detail = "；".join(
            f"样本号“{sno}”重复（第 {first} 行、第 {again} 行）"
            for sno, first, again in duplicates
        )
        raise WorkbookError(f"唯一样本号重复，文件被拒收：{detail}。")
    if not rows:
        raise WorkbookError("表头之下没有任何数据行。")
    return rows


def validate_batch_continuity(rows: List[RowRecord]) -> None:
    """业务规则校验：同一批号的行必须连续，且原始站码相同。"""
    # 批号连续性：同一个批号在表中只能出现在一个连续区间
    segments: Dict[str, List[int]] = {}
    prev_batch = None
    for rec in rows:
        if rec.batch != prev_batch:
            segments.setdefault(rec.batch, []).append(rec.row_index)
            prev_batch = rec.batch
    split = {b: rs for b, rs in segments.items() if len(rs) > 1}
    if split:
        detail = "；".join(
            f"批号“{b}”被其他批次打断，涉及起始行 {rs}" for b, rs in split.items()
        )
        raise WorkbookError(f"同批行必须连续：{detail}。")

    # 同批原始站码必须一致（合并格展开后天然一致，这里兜底校验）
    batch_stations: Dict[str, set] = {}
    for rec in rows:
        batch_stations.setdefault(rec.batch, set()).add(rec.original_station)
    bad = {b: sts for b, sts in batch_stations.items() if len(sts) > 1}
    if bad:
        detail = "；".join(
            f"批号“{b}”存在多个原始站码 {sorted(sts)}" for b, sts in bad.items()
        )
        raise WorkbookError(f"同批行原始站码必须相同：{detail}。")


def parse_upload(path: str, sheet_name: Optional[str] = None) -> ParsedWorkbook:
    """上传解析主流程：识别表头 -> 展开站码 -> 逐行校验。

    使用 data_only=False 保留公式原文。
    """
    try:
        wb = load_workbook(path, data_only=False)
    except Exception as exc:  # openpyxl 读坏文件会抛各类异常
        raise WorkbookError(f"无法打开 XLSX 文件：{exc}") from exc

    ws = wb[sheet_name] if sheet_name else wb.active
    spec = detect_columns(ws)
    station_merges = unmerge_station_cells(ws, spec)
    rows = parse_rows(ws, spec)
    validate_batch_continuity(rows)

    parsed = ParsedWorkbook(
        path=path,
        sheet_name=ws.title,
        spec=spec,
        rows=rows,
        station_merges=station_merges,
    )
    parsed.workbook = wb  # 临时挂接，供调用方决定是否保存
    return parsed


# --------------------------------------------------------------------------- #
# 更正应用
# --------------------------------------------------------------------------- #
def latest_batch_corrections(rows: List[RowRecord], batch_corrections: List[dict]):
    """按批号取最新一条整批更正（id 越大越新）。"""
    latest: Dict[str, dict] = {}
    for corr in sorted(batch_corrections, key=lambda c: c["id"]):
        latest[corr["batch_no"]] = corr
    return latest


def latest_sample_corrections(sample_corrections: List[dict]):
    """按样本号取最新一条样本更正（含撤销记录）。"""
    latest: Dict[str, dict] = {}
    for corr in sorted(sample_corrections, key=lambda c: c["id"]):
        latest[corr["sample_no"]] = corr
    return latest


def apply_corrections(
    rows: List[RowRecord],
    batch_corrections: List[dict],
    sample_corrections: List[dict],
) -> List[RowRecord]:
    """计算每行的现行有效站码。

    优先级：
      1. 样本更正（action=correct）——即使值等于当时批站码，也锁定该值，
         之后整批改站不影响该行；
      2. 样本撤销（action=revoke）或无样本更正 -> 采用现行批站码：
         该批最新整批改站的目标站码，没有整批改站则用原始站码。
    """
    batch_latest = latest_batch_corrections(rows, batch_corrections)
    sample_latest = latest_sample_corrections(sample_corrections)

    for rec in rows:
        rec.batch_correction = batch_latest.get(rec.batch)
        current_batch_station = (
            rec.batch_correction["new_station"]
            if rec.batch_correction
            else rec.original_station
        )
        sample_corr = sample_latest.get(rec.sample_no)
        rec.sample_correction = sample_corr

        if sample_corr and sample_corr["action"] == "correct":
            rec.effective_station = sample_corr["new_station"]
            rec.source = "sample"
        else:
            rec.effective_station = current_batch_station
            rec.source = "batch" if rec.batch_correction else "original"
            if sample_corr and sample_corr["action"] == "revoke":
                rec.source = "revoked->batch" if rec.batch_correction else "revoked->original"
    return rows


# --------------------------------------------------------------------------- #
# 导出：重新合并站码并保存工作簿
# --------------------------------------------------------------------------- #
def merge_effective_stations(ws, spec: ColumnSpec, rows: List[RowRecord]) -> List[Tuple[int, int]]:
    """把同批、连续、同有效站码的行重新纵向合并站码。

    跨批即使站码相同也分段，不合并。返回合并区间列表。
    """
    station_col = spec.station_col
    merged: List[Tuple[int, int]] = []
    if not rows:
        return merged

    def flush(start_rec: RowRecord, end_row: int, value: str):
        ws.cell(row=start_rec.row_index, column=station_col).value = value
        if end_row > start_rec.row_index:
            ws.merge_cells(
                start_row=start_rec.row_index,
                start_column=station_col,
                end_row=end_row,
                end_column=station_col,
            )
            merged.append((start_rec.row_index, end_row))

    seg_start = rows[0]
    seg_value = rows[0].effective_station
    seg_batch = rows[0].batch
    prev_row = rows[0]

    for rec in rows[1:]:
        same_segment = (
            rec.batch == seg_batch
            and rec.effective_station == seg_value
            and rec.row_index == prev_row.row_index + 1
        )
        if not same_segment:
            flush(seg_start, prev_row.row_index, seg_value)
            seg_start = rec
            seg_value = rec.effective_station
            seg_batch = rec.batch
        prev_row = rec
    flush(seg_start, prev_row.row_index, seg_value)
    return merged

def build_corrected_workbook(
    original_path: str,
    spec: ColumnSpec,
    rows: List[RowRecord],
    sheet_name: str,
    output_path: str,
) -> List[Tuple[int, int]]:
    """基于原表生成更正工作簿。

    - 从原文件重新加载（data_only=False），保留行序、其他列的值与公式、样式；
    - 先展开原始站码合并，写入有效站码，再按新分段重新合并；
    - 其他列与其他工作表原样保留。
    """
    wb = load_workbook(original_path, data_only=False)
    ws = wb[sheet_name]

    # 展开站码列原有合并（其他列合并保持不动）
    unmerge_station_cells(ws, spec)

    station_col = spec.station_col
    # 先逐行写值，保证分段内每一行都有正确值
    for rec in rows:
        ws.cell(row=rec.row_index, column=station_col).value = rec.effective_station

    merged = merge_effective_stations(ws, spec, rows)
    try:
        wb.save(output_path)
    except Exception as exc:
        raise WorkbookError(f"更正工作簿生成失败：{exc}") from exc
    return merged
