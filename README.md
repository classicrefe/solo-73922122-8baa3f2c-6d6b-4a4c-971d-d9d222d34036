# 水质送检表入仓更正平台

用 Flask + openpyxl + SQLite 实现的送检表站码更正网页：上传单表 XLSX，
自动展开纵向合并的采样站码，支持**整批改站**与**按样本更正/撤销**，
下载时再按规则重新合并站码生成更正工作簿。

## 业务规则

1. **表结构**：表头含“采样站码 / 批号 / 唯一样本号”；批号、样本号逐行填写；
   同批行连续，站码以纵向合并格共用一个值。
2. **展开**：上传后站码列合并格自动展开，合并左上角值填入该段每一行；其他列不动。
3. **两类更正分表存储**（SQLite）：
   - 整批更正 `batch_corrections`：对批号下所有行生效；
   - 样本更正 `sample_corrections`：按唯一样本号生效，动作分 `correct` / `revoke`。
4. **优先级**：样本更正 > 整批改站 > 原始站码。
   - 即使样本更正填写的就是当时的批站码，也会**锁定**该值，之后整批改站不影响该行；
   - 撤销样本更正后，回退采用**现行批站码**（有整批改站取最新目标值，否则取原始站码）。
5. **下载重新合并**：同批、连续、同有效站码的行重新纵向合并站码；
   **跨批即使站码相同也分开**；行序、其他列的值与公式原样保留。
6. **拒收与容错**：
   - 唯一样本号重复 → 拒收，并指出重复样本号所在的行号；
   - 同批不连续、同批原始站码不一致、缺表头/缺必填值 → 拒收并说明原因；
   - 拒收时原表移动到 `data/rejected/` 保留，提示修正后重试；
   - 更正工作簿生成失败时，原表与全部更正记录保留，页面提示重试。

## 目录结构

```
app/
  app.py            # Flask 路由与页面流程
  excel_service.py  # openpyxl 核心逻辑（展开/校验/优先级/重新合并/导出）
  db.py             # SQLite 建表与两类更正存取
  templates/        # index.html / detail.html / base.html
tests/              # 17 个 unittest 用例（核心规则 + 网页端到端）
scripts/
  make_demo_xlsx.py # 生成演示 XLSX
  smoke_test.py     # 纯标准库 HTTP 全流程冒烟脚本
wsgi.py             # 应用入口
Dockerfile
docker-compose.yml
requirements.txt
```

## 用 Docker Compose 启动（推荐）

```bash
docker compose up --build
# 浏览器打开 http://localhost:5000
```

数据（上传原表、导出文件、SQLite、拒收文件）持久化在宿主机 `./data` 目录。

## 本地直接运行

```bash
python3 -m pip install -r requirements.txt
python3 wsgi.py          # http://127.0.0.1:5000
```

## 网页演示路径

1. 首页上传 `water_demo.xlsx`（用 `python3 scripts/make_demo_xlsx.py /tmp/demo` 生成）；
2. 详情页左侧对批号 `B2401` 整批改站为 `S07`；
3. 右侧对样本 `SMP-002` 做样本更正为 `S01`（等于原始批站码，验证锁定）；
4. 再把 `B2401` 整批改到 `S08`，可见 `SMP-002` 仍锁定 `S01`，同批其他行为 `S08`；
5. 撤销 `SMP-002` 的样本更正，该行回退现行批站码 `S08`；
6. 点击“下载更正工作簿”，核验站码合并分段与公式保留；
7. 上传 `water_duplicate.xlsx`，看到重复样本号行号提示与拒收保留说明。

## 测试与冒烟

```bash
# 单元 + 网页端到端（17 个用例）
python3 -m unittest discover -s tests -v

# 真实 HTTP 冒烟（先启动服务）
python3 scripts/make_demo_xlsx.py /tmp/demo
python3 wsgi.py &
python3 scripts/smoke_test.py http://127.0.0.1:5000
```

## 设计说明

- 导出时从**原始上传文件**重新加载工作簿（`data_only=False`）再套用现行有效站码，
  因此公式以原文保留，其他工作表、列宽样式、非站码列的合并格均不受影响。
- 每次更正都追加写库、不删除历史，`old_station` 留存更正前快照，便于审计。
- 表头采用关键字包含匹配（如“采样站码/站码”“批号/批次号”“唯一样本号/样品编号”），
  三列表头须在同一行，否则拒收，避免列错位。
