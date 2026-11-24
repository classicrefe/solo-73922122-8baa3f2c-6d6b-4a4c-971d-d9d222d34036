# -*- coding: utf-8 -*-
"""Flask 端到端测试：上传 -> 整批改站 -> 样本更正/撤销 -> 下载 -> 拒收与保留。"""
import io
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook

from app.app import create_app
from tests.fixtures import (
    build_demo_workbook,
    build_duplicate_workbook,
)


class WebFlowTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        app = create_app()
        app.config["DATA_DIR"] = self.tmp.name
        app.config["UPLOAD_DIR"] = os.path.join(self.tmp.name, "uploads")
        app.config["EXPORT_DIR"] = os.path.join(self.tmp.name, "exports")
        app.config["DATABASE_PATH"] = os.path.join(self.tmp.name, "test.sqlite3")
        app.config["WTF_CSRF_ENABLED"] = False
        # create_app 已按默认目录建好文件夹，这里按测试临时目录重建
        os.makedirs(app.config["UPLOAD_DIR"], exist_ok=True)
        os.makedirs(app.config["EXPORT_DIR"], exist_ok=True)
        # 重新指向临时库并建表
        from app import db
        db.init_db(app)
        self.client = app.test_client()
        self.app = app

        self.src = os.path.join(self.tmp.name, "demo.xlsx")
        build_demo_workbook(self.src)
        self.dup = os.path.join(self.tmp.name, "dup.xlsx")
        build_duplicate_workbook(self.dup)

    def tearDown(self):
        self.tmp.cleanup()

    def _upload(self, path):
        with open(path, "rb") as fh:
            data = {"file": (fh, os.path.basename(path))}
            return self.client.post(
                "/upload", data=data, content_type="multipart/form-data",
                follow_redirects=True,
            )

    def test_full_correction_flow_and_download(self):
        # 1) 上传
        resp = self._upload(self.src)
        self.assertEqual(resp.status_code, 200)
        self.assertIn("上传成功".encode(), resp.data)
        self.assertIn("展开站码合并格 3 处".encode(), resp.data)

        # 2) 整批 B2401 改站 S07
        resp = self.client.post(
            "/workbooks/1/batch-correct",
            data={"batch_no": "B2401", "new_station": "S07", "operator": "张三"},
            follow_redirects=True,
        )
        self.assertIn("已登记整批改站".encode(), resp.data)
        # 详情页应显示 S07 与“整批改站”标签
        self.assertIn("S07".encode(), resp.data)

        # 3) 样本 SMP-002 更正为 S01（=原始批站码），验证锁定语义
        resp = self.client.post(
            "/workbooks/1/sample-correct",
            data={"sample_no": "SMP-002", "new_station": "S01", "operator": "李四"},
            follow_redirects=True,
        )
        self.assertIn("已登记样本更正".encode(), resp.data)
        self.assertIn("样本更正".encode(), resp.data)

        # 4) 再整批 B2401 改站 S08，SMP-002 仍应为 S01
        self.client.post(
            "/workbooks/1/batch-correct",
            data={"batch_no": "B2401", "new_station": "S08"},
            follow_redirects=True,
        )
        resp = self.client.get("/workbooks/1")
        self.assertIn("S08".encode(), resp.data)
        self.assertIn("S01".encode(), resp.data)

        # 5) 撤销 SMP-002，回退现行批站码 S08
        resp = self.client.post(
            "/workbooks/1/sample-revoke",
            data={"sample_no": "SMP-002"},
            follow_redirects=True,
        )
        self.assertIn("已撤销样本".encode(), resp.data)

        # 6) 下载并核验工作簿
        resp = self.client.get("/workbooks/1/download")
        self.assertEqual(resp.status_code, 200)
        self.assertTrue(
            resp.headers["Content-Disposition"].startswith("attachment;")
        )
        out_path = os.path.join(self.tmp.name, "downloaded.xlsx")
        with open(out_path, "wb") as fh:
            fh.write(resp.data)
        wb = load_workbook(out_path)
        ws = wb.active
        # B2401 三行现行站码 S08 -> B2:B4 合并
        self.assertIn("B2:B4", [str(r) for r in ws.merged_cells.ranges])
        self.assertEqual(ws["B2"].value, "S08")
        # B2402 仍是 S02（B5:B6），B2403 仍是 S01（B7:B8）
        self.assertEqual(ws["B5"].value, "S02")
        self.assertEqual(ws["B7"].value, "S01")
        # 公式保留
        self.assertEqual(ws["G3"].value, "=F3*2")
        # 其他列值保留、行序保留
        self.assertEqual(ws["D2"].value, "SMP-001")
        self.assertEqual(ws["D8"].value, "SMP-007")

        # 7) SQLite 两类更正分别入表
        import sqlite3
        conn = sqlite3.connect(self.app.config["DATABASE_PATH"])
        batch_rows = conn.execute(
            "SELECT batch_no, new_station FROM batch_corrections ORDER BY id"
        ).fetchall()
        sample_rows = conn.execute(
            "SELECT sample_no, action, new_station FROM sample_corrections ORDER BY id"
        ).fetchall()
        conn.close()
        self.assertEqual(
            batch_rows, [("B2401", "S07"), ("B2401", "S08")]
        )
        self.assertEqual(
            sample_rows,
            [("SMP-002", "correct", "S01"), ("SMP-002", "revoke", None)],
        )

    def test_duplicate_upload_rejected_and_file_kept(self):
        resp = self._upload(self.dup)
        self.assertEqual(resp.status_code, 200)
        self.assertIn("唯一样本号重复".encode(), resp.data)
        self.assertIn("SMP-002".encode(), resp.data)
        # 原表保留在 rejected 目录
        rejected = os.listdir(os.path.join(self.tmp.name, "rejected"))
        self.assertEqual(len(rejected), 1)
        self.assertTrue(rejected[0].endswith("dup.xlsx"))
        # 数据库没有登记任何工作簿
        import sqlite3
        conn = sqlite3.connect(self.app.config["DATABASE_PATH"])
        count = conn.execute("SELECT COUNT(*) FROM workbooks").fetchone()[0]
        conn.close()
        self.assertEqual(count, 0)

    def test_unknown_sample_and_batch_rejected(self):
        self._upload(self.src)
        resp = self.client.post(
            "/workbooks/1/sample-correct",
            data={"sample_no": "NO-SUCH", "new_station": "S99"},
            follow_redirects=True,
        )
        self.assertIn("不存在".encode(), resp.data)
        resp = self.client.post(
            "/workbooks/1/batch-correct",
            data={"batch_no": "B9999", "new_station": "S99"},
            follow_redirects=True,
        )
        self.assertIn("不存在".encode(), resp.data)

    def test_non_xlsx_rejected(self):
        resp = self.client.post(
            "/upload",
            data={"file": (io.BytesIO(b"abc"), "note.txt")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertIn(".xlsx".encode(), resp.data)

    def test_download_failure_keeps_records(self):
        """模拟导出阶段失败：应提示重试，且原表与更正记录均保留。"""
        self._upload(self.src)
        self.client.post(
            "/workbooks/1/batch-correct",
            data={"batch_no": "B2401", "new_station": "S07"},
            follow_redirects=True,
        )

        # monkeypatch：让工作簿保存环节抛错，模拟生成失败
        import app.app as app_module
        from app.excel_service import WorkbookError

        def boom(**kwargs):
            raise WorkbookError("模拟磁盘写入失败")

        original = app_module.build_corrected_workbook
        app_module.build_corrected_workbook = boom
        try:
            resp = self.client.get("/workbooks/1/download", follow_redirects=True)
        finally:
            app_module.build_corrected_workbook = original

        self.assertIn("生成失败".encode(), resp.data)
        self.assertIn("重试".encode(), resp.data)
        # 原表仍在
        uploads = os.listdir(self.app.config["UPLOAD_DIR"])
        self.assertTrue(any(f.endswith("demo.xlsx") for f in uploads))
        # 更正记录仍可查
        import sqlite3
        conn = sqlite3.connect(self.app.config["DATABASE_PATH"])
        count = conn.execute("SELECT COUNT(*) FROM batch_corrections").fetchone()[0]
        conn.close()
        self.assertEqual(count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
