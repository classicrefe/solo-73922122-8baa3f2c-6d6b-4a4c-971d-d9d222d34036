# -*- coding: utf-8 -*-
"""入仓更正页 Flask 应用：上传、整批改站、样本更正/撤销、下载更正工作簿。"""
from __future__ import annotations

import io
import os
from typing import Dict, List

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

from . import db, excel_service  # noqa: F401  (excel_service 保留供扩展/测试引用)
from .excel_service import (
    WorkbookError,
    apply_corrections,
    build_corrected_workbook,
    parse_upload,
)


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")
    base_dir = os.environ.get(
        "DATA_DIR", os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
    )
    app.config["DATA_DIR"] = base_dir
    app.config["UPLOAD_DIR"] = os.path.join(base_dir, "uploads")
    app.config["EXPORT_DIR"] = os.path.join(base_dir, "exports")
    app.config["DATABASE_PATH"] = os.path.join(base_dir, "corrections.sqlite3")
    app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB
    os.makedirs(app.config["UPLOAD_DIR"], exist_ok=True)
    os.makedirs(app.config["EXPORT_DIR"], exist_ok=True)

    db.init_db(app)
    app.teardown_appcontext(db.close_db)

    # ------------------------------------------------------------------ #
    # 数据装配辅助
    # ------------------------------------------------------------------ #
    def _load_rows(workbook_id: int):
        """读取原表并套用全部更正，返回 (row, spec, sheet_name, stored_path)。"""
        wb_row = db.get_workbook(workbook_id)
        if wb_row is None:
            abort(404)
        stored_path = wb_row["stored_path"]
        if not os.path.exists(stored_path):
            raise WorkbookError(
                "原始上传文件已丢失，无法重建数据；请重新上传（更正记录仍保留）。"
            )
        parsed = parse_upload(stored_path, wb_row["sheet_name"])
        batch_corr = [dict(r) for r in db.list_batch_corrections(workbook_id)]
        sample_corr = [dict(r) for r in db.list_sample_corrections(workbook_id)]
        apply_corrections(parsed.rows, batch_corr, sample_corr)
        return wb_row, parsed

    def _batch_map(rows) -> Dict[str, List]:
        """按批号分组（保持表内出现顺序）。"""
        groups: Dict[str, List] = {}
        for rec in rows:
            groups.setdefault(rec.batch, []).append(rec)
        return groups

    # ------------------------------------------------------------------ #
    # 页面：首页 / 工作簿列表
    # ------------------------------------------------------------------ #
    @app.route("/")
    def index():
        workbooks = db.list_workbooks()
        return render_template("index.html", workbooks=workbooks)

    # ------------------------------------------------------------------ #
    # 上传
    # ------------------------------------------------------------------ #
    @app.route("/upload", methods=["POST"])
    def upload():
        file = request.files.get("file")
        if file is None or not file.filename:
            flash("请选择要上传的 XLSX 文件。", "error")
            return redirect(url_for("index"))
        if not file.filename.lower().endswith(".xlsx"):
            flash("仅支持 .xlsx 单表工作簿（旧版 .xls 请先另存为 .xlsx）。", "error")
            return redirect(url_for("index"))

        # 先落盘再解析：解析失败时原表保留，便于排查与重试
        safe_name = file.filename.replace("/", "_").replace("\\", "_")
        tmp_name = f"pending_{os.getpid()}_{safe_name}"
        tmp_path = os.path.join(app.config["UPLOAD_DIR"], tmp_name)
        file.save(tmp_path)

        try:
            parsed = parse_upload(tmp_path)
            # 先插入登记记录拿到 id，再把原表重命名为与 id 对应的正式文件名
            placeholder_path = os.path.join(
                app.config["UPLOAD_DIR"], f"wb_pending_{os.getpid()}_{safe_name}"
            )
            workbook_id = db.insert_workbook(
                original_name=safe_name,
                stored_path=placeholder_path,
                sheet_name=parsed.sheet_name,
                spec=parsed.spec,
                row_count=len(parsed.rows),
            )
            final_path = os.path.join(
                app.config["UPLOAD_DIR"], f"wb_{workbook_id}_{safe_name}"
            )
            os.replace(tmp_path, final_path)
            conn = db.get_db()
            conn.execute(
                "UPDATE workbooks SET stored_path = ? WHERE id = ?",
                (final_path, workbook_id),
            )
            conn.commit()
            flash(
                f"上传成功：{len(parsed.rows)} 行，"
                f"展开站码合并格 {len(parsed.station_merges)} 处。",
                "success",
            )
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))
        except WorkbookError as exc:
            # 失败保留原表：移动到保留目录而不是删除
            keep_dir = os.path.join(app.config["DATA_DIR"], "rejected")
            os.makedirs(keep_dir, exist_ok=True)
            keep_path = os.path.join(keep_dir, tmp_name)
            os.replace(tmp_path, keep_path)
            flash(f"上传被拒收：{exc} 原表已保留于 rejected/，可修正后重试。", "error")
            return redirect(url_for("index"))
        except Exception as exc:  # 兜底：同样保留原表
            keep_dir = os.path.join(app.config["DATA_DIR"], "rejected")
            os.makedirs(keep_dir, exist_ok=True)
            keep_path = os.path.join(keep_dir, tmp_name)
            if os.path.exists(tmp_path):
                os.replace(tmp_path, keep_path)
            app.logger.exception("上传处理异常")
            flash(
                f"上传处理发生异常：{exc}。原表已保留于 rejected/，请修正后重试。",
                "error",
            )
            return redirect(url_for("index"))

    # ------------------------------------------------------------------ #
    # 工作簿详情：按批展示 + 更正入口
    # ------------------------------------------------------------------ #
    @app.route("/workbooks/<int:workbook_id>")
    def workbook_detail(workbook_id: int):
        try:
            wb_row, parsed = _load_rows(workbook_id)
        except WorkbookError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))
        groups = _batch_map(parsed.rows)
        batch_corr = db.list_batch_corrections(workbook_id)
        sample_corr = db.list_sample_corrections(workbook_id)
        return render_template(
            "detail.html",
            wb=wb_row,
            spec=parsed.spec,
            groups=groups,
            batch_corrections=batch_corr,
            sample_corrections=sample_corr,
        )

    # ------------------------------------------------------------------ #
    # 整批改站
    # ------------------------------------------------------------------ #
    @app.route("/workbooks/<int:workbook_id>/batch-correct", methods=["POST"])
    def batch_correct(workbook_id: int):
        batch_no = (request.form.get("batch_no") or "").strip()
        new_station = (request.form.get("new_station") or "").strip()
        operator = (request.form.get("operator") or "").strip()
        if not batch_no or not new_station:
            flash("批号和新站码均不能为空。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        try:
            _, parsed = _load_rows(workbook_id)
        except WorkbookError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        target = [r for r in parsed.rows if r.batch == batch_no]
        if not target:
            flash(f"原表中不存在批号“{batch_no}”。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        old_station = target[0].original_station
        if new_station == old_station and not target[0].batch_correction:
            flash("新站码与该批原始站码相同，无需更正。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        db.insert_batch_correction(
            workbook_id, batch_no, old_station, new_station, operator
        )
        flash(
            f"已登记整批改站：批号 {batch_no}，"
            f"{old_station} → {new_station}（共 {len(target)} 行）。",
            "success",
        )
        return redirect(url_for("workbook_detail", workbook_id=workbook_id))

    # ------------------------------------------------------------------ #
    # 样本更正 / 撤销
    # ------------------------------------------------------------------ #
    @app.route("/workbooks/<int:workbook_id>/sample-correct", methods=["POST"])
    def sample_correct(workbook_id: int):
        sample_no = (request.form.get("sample_no") or "").strip()
        new_station = (request.form.get("new_station") or "").strip()
        operator = (request.form.get("operator") or "").strip()
        if not sample_no or not new_station:
            flash("样本号和新站码均不能为空。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        try:
            _, parsed = _load_rows(workbook_id)
        except WorkbookError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        target = next((r for r in parsed.rows if r.sample_no == sample_no), None)
        if target is None:
            flash(f"原表中不存在样本号“{sample_no}”。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        # old_station 记录该样本当前有效站码，保留更正前快照
        old_station = target.effective_station
        db.insert_sample_correction(
            workbook_id,
            sample_no,
            target.batch,
            "correct",
            old_station,
            new_station,
            operator,
        )
        flash(
            f"已登记样本更正：样本 {sample_no}，{old_station} → {new_station}。"
            "该更正优先于整批改站。",
            "success",
        )
        return redirect(url_for("workbook_detail", workbook_id=workbook_id))

    @app.route("/workbooks/<int:workbook_id>/sample-revoke", methods=["POST"])
    def sample_revoke(workbook_id: int):
        sample_no = (request.form.get("sample_no") or "").strip()
        operator = (request.form.get("operator") or "").strip()
        if not sample_no:
            flash("样本号不能为空。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        try:
            _, parsed = _load_rows(workbook_id)
        except WorkbookError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        target = next((r for r in parsed.rows if r.sample_no == sample_no), None)
        if target is None:
            flash(f"原表中不存在样本号“{sample_no}”。", "error")
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        # 撤销后该行将采用的站码：有整批改站取最新目标值，否则取原始站码
        fallback_station = (
            target.batch_correction["new_station"]
            if target.batch_correction
            else target.original_station
        )
        old_station = target.effective_station
        db.insert_sample_correction(
            workbook_id,
            sample_no,
            target.batch,
            "revoke",
            old_station,
            None,
            operator,
        )
        flash(
            f"已撤销样本 {sample_no} 的单独更正，恢复采用现行批站码"
            f"（当前为 {fallback_station}）。",
            "success",
        )
        return redirect(url_for("workbook_detail", workbook_id=workbook_id))

    # ------------------------------------------------------------------ #
    # 下载更正工作簿
    # ------------------------------------------------------------------ #
    @app.route("/workbooks/<int:workbook_id>/download")
    def download(workbook_id: int):
        try:
            wb_row, parsed = _load_rows(workbook_id)
        except WorkbookError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        base = os.path.splitext(wb_row["original_name"])[0]
        out_name = f"{base}_corrected.xlsx"
        out_path = os.path.join(app.config["EXPORT_DIR"], f"wb_{workbook_id}_{out_name}")
        try:
            merged = build_corrected_workbook(
                original_path=wb_row["stored_path"],
                spec=parsed.spec,
                rows=parsed.rows,
                sheet_name=parsed.sheet_name,
                output_path=out_path,
            )
        except WorkbookError as exc:
            # 生成失败：原表与更正记录均保留，提示重试
            flash(
                f"更正工作簿生成失败：{exc}。原表与全部更正记录均已保留，请重试下载。",
                "error",
            )
            return redirect(url_for("workbook_detail", workbook_id=workbook_id))

        with open(out_path, "rb") as fh:
            data = fh.read()
        flash(
            f"更正工作簿已生成：{len(parsed.rows)} 行，重新合并站码 {len(merged)} 处。",
            "success",
        )
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    # 模板里用到的小工具
    @app.template_filter("station_source")
    def _station_source(source: str) -> str:
        return {
            "original": "原始站码",
            "batch": "整批改站",
            "sample": "样本更正",
            "revoked->batch": "已撤销→整批",
            "revoked->original": "已撤销→原始",
        }.get(source, source)

    @app.template_filter("action_label")
    def _action_label(action: str) -> str:
        return {"correct": "更正", "revoke": "撤销"}.get(action, action)

    return app
