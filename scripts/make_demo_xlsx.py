# -*- coding: utf-8 -*-
"""端到端冒烟：生成演示 XLSX，供 smoke.sh 上传使用。

用法：python3 scripts/make_demo_xlsx.py <输出路径>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.fixtures import build_demo_workbook, build_duplicate_workbook

if __name__ == "__main__":
    out_dir = sys.argv[1] if len(sys.argv) > 1 else "/tmp/wq-demo"
    os.makedirs(out_dir, exist_ok=True)
    demo = os.path.join(out_dir, "water_demo.xlsx")
    dup = os.path.join(out_dir, "water_duplicate.xlsx")
    build_demo_workbook(demo)
    build_duplicate_workbook(dup)
    print(demo)
    print(dup)
