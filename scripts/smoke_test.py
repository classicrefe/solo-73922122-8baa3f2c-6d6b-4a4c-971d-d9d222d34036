# -*- coding: utf-8 -*-
"""网页全流程冒烟脚本（不依赖测试框架，直接打 HTTP）。

步骤：
  1. 生成演示 XLSX 与重复样本号 XLSX（独立临时目录）；
  2. 上传演示表，跟随跳转抓取 workbook id；
  3. 整批改站 -> 样本更正 -> 再次整批改站 -> 撤销样本更正；
  4. 下载更正工作簿并用 openpyxl 核验合并区间、公式、行序；
  5. 上传重复样本号文件，验证拒收提示与 rejected/ 保留。

用法：python3 scripts/smoke_test.py [base_url]
"""
import io
import os
import re
import sys
import tempfile
import urllib.parse
import urllib.request
import uuid

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook

from tests.fixtures import build_demo_workbook, build_duplicate_workbook


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """禁止自动跟随重定向：302 以 HTTPError 抛出，由调用方手动跟随，
    这样才能保住 302 响应上的 session cookie（flash 消息依赖它）。"""

    def redirect_request(self, *args, **kwargs):
        return None


class Client:
    def __init__(self, base):
        self.base = base.rstrip("/")
        # flash 消息依赖 session cookie，需要 cookie 罐跨请求保持
        self.cookies = {}
        self.opener = urllib.request.build_opener(_NoRedirect)

    def _request(self, path, data=None, files=None):
        url = self.base + path
        headers = {}
        if self.cookies:
            headers["Cookie"] = "; ".join(
                f"{k}={v}" for k, v in self.cookies.items()
            )
        if files is not None:
            boundary = "----smoke" + uuid.uuid4().hex
            body = io.BytesIO()
            for name, value in (data or {}).items():
                body.write(f"--{boundary}\r\n".encode())
                body.write(
                    f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode()
                )
                body.write(f"{value}\r\n".encode())
            for name, (filename, content) in files.items():
                body.write(f"--{boundary}\r\n".encode())
                body.write(
                    f'Content-Disposition: form-data; name="{name}"; '
                    f'filename="{filename}"\r\n'.encode()
                )
                body.write(b"Content-Type: application/octet-stream\r\n\r\n")
                body.write(content)
                body.write(b"\r\n")
            body.write(f"--{boundary}--\r\n".encode())
            req = urllib.request.Request(
                url, data=body.getvalue(),
                headers={
                    "Content-Type": f"multipart/form-data; boundary={boundary}",
                    **headers,
                },
            )
        elif data is not None:
            encoded = urllib.parse.urlencode(data).encode()
            req = urllib.request.Request(url, data=encoded, headers=headers)
        else:
            req = urllib.request.Request(url, headers=headers)
        try:
            resp = self.opener.open(req)
            self._store_cookies(resp.headers)
            return resp.status, resp.read(), resp.headers
        except urllib.error.HTTPError as exc:
            # 302/301/303：先存 cookie（session 里带着 flash），再手动跟随
            self._store_cookies(exc.headers)
            if exc.code in (301, 302, 303) and "Location" in exc.headers:
                loc = exc.headers["Location"]
                # 统一裁成路径，经 self.base 发起请求，避免绝对地址解析差异
                parsed_loc = urllib.parse.urlparse(loc)
                path = parsed_loc.path or "/"
                if parsed_loc.query:
                    path += "?" + parsed_loc.query
                return self.get(path)
            return exc.code, exc.read(), exc.headers

    def _store_cookies(self, resp_headers):
        for cookie in resp_headers.get_all("Set-Cookie", []) or []:
            pair = cookie.split(";", 1)[0]
            if "=" in pair:
                key, value = pair.split("=", 1)
                self.cookies[key.strip()] = value.strip()

    def get(self, path):
        return self._request(path)

    def post(self, path, data=None, files=None):
        return self._request(path, data=data, files=files)


def must(cond, msg):
    if not cond:
        print(f"  [FAIL] {msg}")
        sys.exit(1)
    print(f"  [OK]   {msg}")


def main():
    base = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:5000"
    client = Client(base)
    workdir = tempfile.mkdtemp(prefix="wq-smoke-")
    demo = os.path.join(workdir, "water_demo.xlsx")
    dup = os.path.join(workdir, "water_duplicate.xlsx")
    build_demo_workbook(demo)
    build_duplicate_workbook(dup)
    print(f"演示文件目录：{workdir}")

    print("1) 上传演示表")
    with open(demo, "rb") as fh:
        status, body, _ = client.post(
            "/upload", files={"file": ("water_demo.xlsx", fh.read())}
        )
    text = body.decode("utf-8")
    must(status == 200, f"上传返回 200（实际 {status}）")
    must("上传成功" in text, "页面提示上传成功")
    m = re.search(r"/workbooks/(\d+)", text)
    must(m is not None, "跳转到工作簿详情页")
    wb_id = m.group(1)
    detail = f"/workbooks/{wb_id}"

    print("2) 整批 B2401 改站 S07")
    _, body, _ = client.post(
        detail + "/batch-correct",
        data={"batch_no": "B2401", "new_station": "S07", "operator": "复核员甲"},
    )
    must("已登记整批改站" in body.decode(), "整批改站登记成功")

    print("3) 样本 SMP-002 更正为 S01（当时批站码的原值，验证锁定）")
    _, body, _ = client.post(
        detail + "/sample-correct",
        data={"sample_no": "SMP-002", "new_station": "S01", "operator": "复核员乙"},
    )
    must("已登记样本更正" in body.decode(), "样本更正登记成功")

    print("4) 整批 B2401 再次改站 S08，SMP-002 应仍锁定 S01")
    _, body, _ = client.post(
        detail + "/batch-correct",
        data={"batch_no": "B2401", "new_station": "S08"},
    )
    text = body.decode()
    must("S08" in text, "同批其他行显示新批站码 S08")
    must("样本更正" in text and "S01" in text, "SMP-002 仍显示样本更正 S01")

    print("5) 撤销 SMP-002，回退现行批站码 S08")
    _, body, _ = client.post(
        detail + "/sample-revoke", data={"sample_no": "SMP-002"}
    )
    must("已撤销样本" in body.decode(), "撤销成功并提示回退批站码")

    print("6) 下载更正工作簿并核验")
    status, body, headers = client.get(detail + "/download")
    must(status == 200, "下载返回 200")
    must("attachment" in headers.get("Content-Disposition", ""), "响应为附件下载")
    out = os.path.join(workdir, "corrected.xlsx")
    with open(out, "wb") as fh:
        fh.write(body)
    ws = load_workbook(out).active
    merges = sorted(str(r) for r in ws.merged_cells.ranges)
    must(merges == ["B2:B4", "B5:B6", "B7:B8"], f"站码合并区间正确：{merges}")
    must(ws["B2"].value == "S08", "B2401 现行站码 S08（撤销后回退批站码）")
    must(ws["B5"].value == "S02", "B2402 保持 S02")
    must(ws["B7"].value == "S01", "B2403 保持 S01，跨批未与 B2401 合并")
    must(ws["G5"].value == "=F5*2", "其他列公式原样保留（G5 == =F5*2）")
    must(ws["D8"].value == "SMP-007", "行序与样本号原样保留")

    print("7) 重复样本号文件应拒收并指出行号、保留原表")
    with open(dup, "rb") as fh:
        status, body, _ = client.post(
            "/upload", files={"file": ("water_duplicate.xlsx", fh.read())}
        )
    text = body.decode()
    must("唯一样本号重复" in text, "提示样本号重复拒收")
    must("第 3 行" in text and "第 5 行" in text, "指出重复位置（第 3、5 行）")
    must("rejected" in text, "提示原表已保留可重试")

    print("\n全部冒烟检查通过 ✔")
    print(f"更正工作簿样本：{out}")


if __name__ == "__main__":
    main()
