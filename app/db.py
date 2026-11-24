# -*- coding: utf-8 -*-
"""SQLite 持久层：工作簿登记、整批更正、样本更正（含撤销）分表存储。"""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from typing import List, Optional

from flask import current_app, g


SCHEMA = """
CREATE TABLE IF NOT EXISTS workbooks (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    original_name   TEXT NOT NULL,
    stored_path     TEXT NOT NULL,
    sheet_name      TEXT NOT NULL,
    header_row      INTEGER NOT NULL,
    station_col     INTEGER NOT NULL,
    batch_col       INTEGER NOT NULL,
    sample_col      INTEGER NOT NULL,
    headers_json    TEXT NOT NULL,
    row_count       INTEGER NOT NULL,
    uploaded_at     TEXT NOT NULL,
    status          TEXT NOT NULL DEFAULT 'active'
);

CREATE TABLE IF NOT EXISTS batch_corrections (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    workbook_id         INTEGER NOT NULL REFERENCES workbooks(id),
    batch_no            TEXT NOT NULL,
    old_station         TEXT NOT NULL,
    new_station         TEXT NOT NULL,
    operator            TEXT NOT NULL DEFAULT '',
    created_at          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_batch_corr_wb
    ON batch_corrections(workbook_id, id);

CREATE TABLE IF NOT EXISTS sample_corrections (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    workbook_id         INTEGER NOT NULL REFERENCES workbooks(id),
    sample_no           TEXT NOT NULL,
    batch_no            TEXT NOT NULL,
    action              TEXT NOT NULL CHECK (action IN ('correct', 'revoke')),
    old_station         TEXT NOT NULL,
    new_station         TEXT,          -- revoke 时为 NULL
    operator            TEXT NOT NULL DEFAULT '',
    created_at          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sample_corr_wb
    ON sample_corrections(workbook_id, id);
"""


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_db() -> sqlite3.Connection:
    """请求级连接，行按 sqlite3.Row 返回。"""
    if "db" not in g:
        conn = sqlite3.connect(
            current_app.config["DATABASE_PATH"],
            detect_types=sqlite3.PARSE_DECLTYPES,
        )
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        g.db = conn
    return g.db


def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db(app) -> None:
    """应用启动时建表（幂等）。"""
    conn = sqlite3.connect(app.config["DATABASE_PATH"])
    try:
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()


# --------------------------------------------------------------------------- #
# workbooks
# --------------------------------------------------------------------------- #
def insert_workbook(
    original_name: str,
    stored_path: str,
    sheet_name: str,
    spec,
    row_count: int,
) -> int:
    db = get_db()
    cur = db.execute(
        """
        INSERT INTO workbooks (
            original_name, stored_path, sheet_name, header_row,
            station_col, batch_col, sample_col, headers_json,
            row_count, uploaded_at, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active')
        """,
        (
            original_name,
            stored_path,
            sheet_name,
            spec.header_row,
            spec.station_col,
            spec.batch_col,
            spec.sample_col,
            json.dumps(spec.headers, ensure_ascii=False),
            row_count,
            now_str(),
        ),
    )
    db.commit()
    return cur.lastrowid


def list_workbooks() -> List[sqlite3.Row]:
    db = get_db()
    return db.execute(
        "SELECT * FROM workbooks ORDER BY id DESC"
    ).fetchall()


def get_workbook(workbook_id: int) -> Optional[sqlite3.Row]:
    db = get_db()
    return db.execute(
        "SELECT * FROM workbooks WHERE id = ?", (workbook_id,)
    ).fetchone()


# --------------------------------------------------------------------------- #
# 整批更正
# --------------------------------------------------------------------------- #
def insert_batch_correction(
    workbook_id: int,
    batch_no: str,
    old_station: str,
    new_station: str,
    operator: str,
) -> int:
    db = get_db()
    cur = db.execute(
        """
        INSERT INTO batch_corrections
            (workbook_id, batch_no, old_station, new_station, operator, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (workbook_id, batch_no, old_station, new_station, operator, now_str()),
    )
    db.commit()
    return cur.lastrowid


def list_batch_corrections(workbook_id: int) -> List[sqlite3.Row]:
    db = get_db()
    return db.execute(
        """
        SELECT * FROM batch_corrections
        WHERE workbook_id = ? ORDER BY id ASC
        """,
        (workbook_id,),
    ).fetchall()


# --------------------------------------------------------------------------- #
# 样本更正 / 撤销
# --------------------------------------------------------------------------- #
def insert_sample_correction(
    workbook_id: int,
    sample_no: str,
    batch_no: str,
    action: str,
    old_station: str,
    new_station: Optional[str],
    operator: str,
) -> int:
    db = get_db()
    cur = db.execute(
        """
        INSERT INTO sample_corrections
            (workbook_id, sample_no, batch_no, action,
             old_station, new_station, operator, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            workbook_id,
            sample_no,
            batch_no,
            action,
            old_station,
            new_station,
            operator,
            now_str(),
        ),
    )
    db.commit()
    return cur.lastrowid


def list_sample_corrections(workbook_id: int) -> List[sqlite3.Row]:
    db = get_db()
    return db.execute(
        """
        SELECT * FROM sample_corrections
        WHERE workbook_id = ? ORDER BY id ASC
        """,
        (workbook_id,),
    ).fetchall()
