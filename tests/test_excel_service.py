# -*- coding: utf-8 -*-
"""excel_service 单元测试：展开、校验、优先级、撤销、重新合并、公式保留。"""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook

from app.excel_service import (
    WorkbookError,
    apply_corrections,
    build_corrected_workbook,
    parse_upload,
)
from tests.fixtures import (
    HEADERS,
    build_demo_workbook,
    build_duplicate_workbook,
    build_missing_header_workbook,
    build_split_batch_workbook,
)


class ExcelServiceTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.src = os.path.join(self.tmp.name, "demo.xlsx")
        build_demo_workbook(self.src)

    def tearDown(self):
        self.tmp.cleanup()

    def test_unmerge_expands_station_to_every_row(self):
        parsed = parse_upload(self.src)
        self.assertEqual(len(parsed.rows), 7)
        # 三处原始合并格被记录
        self.assertEqual(sorted(parsed.station_merges), [(2, 4), (5, 6), (7, 8)])
        # 展开后每行都有站码
        self.assertEqual(
            [r.original_station for r in parsed.rows],
            ["S01", "S01", "S01", "S02", "S02", "S01", "S01"],
        )
        # 列定位：站码 B(2)、批号 C(3)、样本号 D(4)
        self.assertEqual(
            (parsed.spec.station_col, parsed.spec.batch_col, parsed.spec.sample_col),
            (2, 3, 4),
        )

    def test_duplicate_sample_rejected_with_row_numbers(self):
        dup_path = os.path.join(self.tmp.name, "dup.xlsx")
        build_duplicate_workbook(dup_path)
        with self.assertRaises(WorkbookError) as ctx:
            parse_upload(dup_path)
        msg = str(ctx.exception)
        self.assertIn("SMP-002", msg)
        self.assertIn("第 3 行", msg)
        self.assertIn("第 5 行", msg)

    def test_split_batch_rejected(self):
        p = os.path.join(self.tmp.name, "split.xlsx")
        build_split_batch_workbook(p)
        with self.assertRaises(WorkbookError) as ctx:
            parse_upload(p)
        self.assertIn("连续", str(ctx.exception))

    def test_missing_header_rejected(self):
        p = os.path.join(self.tmp.name, "missing.xlsx")
        build_missing_header_workbook(p)
        with self.assertRaises(WorkbookError) as ctx:
            parse_upload(p)
        self.assertIn("唯一样本号", str(ctx.exception))

    def test_sample_correction_takes_priority_over_later_batch_change(self):
        parsed = parse_upload(self.src)
        # 1) 先做样本更正：SMP-002 -> S09（值与当时批站码 S01 不同）
        corrections_sample = [
            dict(id=1, sample_no="SMP-002", action="correct", new_station="S09"),
        ]
        apply_corrections(parsed.rows, [], corrections_sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-002")
        self.assertEqual(rec.effective_station, "S09")
        self.assertEqual(rec.source, "sample")

        # 2) 之后整批 B2401 改站 S07，样本更正仍优先
        batch = [
            dict(id=2, batch_no="B2401", new_station="S07"),
        ]
        apply_corrections(parsed.rows, batch, corrections_sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-002")
        self.assertEqual(rec.effective_station, "S09")
        # 同批其他行跟随整批
        rec1 = next(r for r in parsed.rows if r.sample_no == "SMP-001")
        self.assertEqual(rec1.effective_station, "S07")
        self.assertEqual(rec1.source, "batch")

    def test_sample_correction_equal_to_then_batch_station_still_locks(self):
        """样本更正填的是当时批站码（S01），后来整批改 S07，样本行仍锁定 S01。"""
        parsed = parse_upload(self.src)
        sample = [
            dict(id=1, sample_no="SMP-003", action="correct", new_station="S01"),
        ]
        batch = [
            dict(id=2, batch_no="B2401", new_station="S07"),
        ]
        apply_corrections(parsed.rows, batch, sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-003")
        self.assertEqual(rec.effective_station, "S01")
        self.assertEqual(rec.source, "sample")

    def test_revoke_falls_back_to_current_batch_station(self):
        parsed = parse_upload(self.src)
        batch = [dict(id=1, batch_no="B2401", new_station="S07")]
        sample = [
            dict(id=2, sample_no="SMP-002", action="correct", new_station="S09"),
            dict(id=3, sample_no="SMP-002", action="revoke", new_station=None),
        ]
        apply_corrections(parsed.rows, batch, sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-002")
        # 撤销后采用现行批站码 S07
        self.assertEqual(rec.effective_station, "S07")
        self.assertIn("revoked", rec.source)

    def test_revoke_without_batch_correction_returns_original(self):
        parsed = parse_upload(self.src)
        sample = [
            dict(id=1, sample_no="SMP-004", action="correct", new_station="S99"),
            dict(id=2, sample_no="SMP-004", action="revoke", new_station=None),
        ]
        apply_corrections(parsed.rows, [], sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-004")
        self.assertEqual(rec.effective_station, "S02")

    def test_export_merges_same_batch_same_station_and_separates_cross_batch(self):
        parsed = parse_upload(self.src)
        # 整批 B2401 -> S01（等于原值，仅为制造场景，实际路由会拦）
        # B2402 -> S01：这样 B2401(S01) 与 B2402(S01) 同站但跨批
        batch = [dict(id=1, batch_no="B2402", new_station="S01")]
        apply_corrections(parsed.rows, batch, [])

        out = os.path.join(self.tmp.name, "out.xlsx")
        merged = build_corrected_workbook(
            self.src, parsed.spec, parsed.rows, parsed.sheet_name, out
        )
        # 全部 7 行有效站码都是 S01，但跨批必须分成 3 段
        self.assertEqual(sorted(merged), [(2, 4), (5, 6), (7, 8)])

        wb = load_workbook(out)
        ws = wb.active
        ranges = sorted(str(r) for r in ws.merged_cells.ranges)
        self.assertEqual(ranges, ["B2:B4", "B5:B6", "B7:B8"])
        # 合并格左上角值
        self.assertEqual(ws["B2"].value, "S01")
        self.assertEqual(ws["B5"].value, "S01")
        self.assertEqual(ws["B7"].value, "S01")
        # 被合并覆盖的单元格值在 openpyxl 中为 None（正常合并行为）
        self.assertIsNone(ws["B3"].value)

    def test_export_split_when_station_differs_within_batch(self):
        """同批内一行样本更正到不同站码：导出时该段必须拆开。"""
        parsed = parse_upload(self.src)
        sample = [
            dict(id=1, sample_no="SMP-002", action="correct", new_station="S09"),
        ]
        apply_corrections(parsed.rows, [], sample)
        out = os.path.join(self.tmp.name, "out2.xlsx")
        merged = build_corrected_workbook(
            self.src, parsed.spec, parsed.rows, parsed.sheet_name, out
        )
        # B2401 三段：行2(S01)、行3(S09)、行4(S01)；行2、行4 不连续不可合并
        self.assertNotIn((2, 4), merged)
        wb = load_workbook(out)
        ws = wb.active
        self.assertEqual(ws["B2"].value, "S01")
        self.assertEqual(ws["B3"].value, "S09")
        self.assertEqual(ws["B4"].value, "S01")
        # B5:B6、B7:B8 仍合并
        self.assertIn("B5:B6", [str(r) for r in ws.merged_cells.ranges])
        self.assertIn("B7:B8", [str(r) for r in ws.merged_cells.ranges])

    def test_export_preserves_row_order_other_values_and_formulas(self):
        parsed = parse_upload(self.src)
        apply_corrections(parsed.rows, [], [])
        out = os.path.join(self.tmp.name, "out3.xlsx")
        build_corrected_workbook(
            self.src, parsed.spec, parsed.rows, parsed.sheet_name, out
        )
        src_wb = load_workbook(self.src)
        out_wb = load_workbook(out)
        src_ws, out_ws = src_wb.active, out_wb.active

        # 行序与其他列值/公式一致（站码列 B 除外，B 列合并区间应与原表一致）
        for r in range(1, 9):
            for c in (1, 3, 4, 5, 6, 7):
                self.assertEqual(
                    out_ws.cell(row=r, column=c).value,
                    src_ws.cell(row=r, column=c).value,
                    f"第{r}行第{c}列发生变化",
                )
        # 公式以原文保留
        self.assertEqual(out_ws["G2"].value, "=F2*2")
        self.assertEqual(out_ws["G8"].value, "=F8*2")
        # 无更正时合并结构与原表相同
        self.assertEqual(
            sorted(str(r) for r in out_ws.merged_cells.ranges),
            sorted(str(r) for r in src_ws.merged_cells.ranges),
        )

    def test_re_correction_overrides_previous_sample_correction(self):
        parsed = parse_upload(self.src)
        sample = [
            dict(id=1, sample_no="SMP-001", action="correct", new_station="S09"),
            dict(id=2, sample_no="SMP-001", action="correct", new_station="S10"),
        ]
        apply_corrections(parsed.rows, [], sample)
        rec = next(r for r in parsed.rows if r.sample_no == "SMP-001")
        self.assertEqual(rec.effective_station, "S10")

    def test_other_column_merges_title_and_style_are_preserved(self):
        """标题行合并、其他列合并、单元格样式在导出后保持不变。"""
        from openpyxl import Workbook
        from openpyxl.styles import Font

        path = os.path.join(self.tmp.name, "styled.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["水质送检表标题"])  # 第 1 行标题
        ws.merge_cells("A1:G1")
        ws["A1"].font = Font(bold=True, size=14)
        ws.append(HEADERS)  # 第 2 行表头
        rows = [
            ("B1", "S1", "S01", 1),
            ("B1", "S2", "S01", 2),
            ("B2", "S3", "S02", 3),
        ]
        for i, (batch, sample, station, val) in enumerate(rows, start=1):
            r = i + 2
            ws.cell(row=r, column=1, value=i)
            ws.cell(row=r, column=2, value=station)
            ws.cell(row=r, column=3, value=batch)
            ws.cell(row=r, column=4, value=sample)
            ws.cell(row=r, column=5, value="pH")
            ws.cell(row=r, column=6, value=val)
            ws.cell(row=r, column=7, value=f"=F{r}*2")
        ws.merge_cells("B3:B4")  # 站码列合并
        ws.merge_cells("E3:E4")  # 其他列（检测项目）合并，必须保留
        header_font = ws.cell(row=2, column=2).font
        ws.cell(row=2, column=2).font = Font(bold=True, color="FF0000")
        wb.save(path)

        parsed = parse_upload(path)
        self.assertEqual(parsed.spec.header_row, 2)  # 表头在第 2 行
        apply_corrections(parsed.rows, [], [])
        out = os.path.join(self.tmp.name, "styled_out.xlsx")
        build_corrected_workbook(
            path, parsed.spec, parsed.rows, parsed.sheet_name, out
        )

        out_ws = load_workbook(out).active
        merges = sorted(str(r) for r in out_ws.merged_cells.ranges)
        # 标题合并、检测项目列合并都保留；站码列按原分段重建
        self.assertIn("A1:G1", merges)
        self.assertIn("E3:E4", merges)
        self.assertIn("B3:B4", merges)
        # 标题与表头样式保留
        self.assertTrue(out_ws["A1"].font.bold)
        self.assertEqual(out_ws["A1"].font.size, 14)
        self.assertEqual(out_ws.cell(row=2, column=2).font.color.rgb, "00FF0000")
        # 标题文本保留
        self.assertEqual(out_ws["A1"].value, "水质送检表标题")


if __name__ == "__main__":
    unittest.main(verbosity=2)
