"""对比 MIYO/OTA 最新两版导出报表，生成可读的差异 Excel。"""
from datetime import datetime
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
TARGETS = [
    ("FYA箱包旗舰店", "京准通快车订单效果明细"),
    ("FYA箱包旗舰店", "京准通全站营销单品推广效果"),
    ("MIYO箱包旗舰店", "京准通快车订单效果明细"),
    ("MIYO箱包旗舰店", "京准通全站营销单品推广效果"),
    ("OTA箱包旗舰店", "京准通快车订单效果明细"),
    ("OTA箱包旗舰店", "京准通全站营销单品推广效果"),
]
DATE_RE = re.compile(r"(20\d{2}-\d{2}-\d{2})")


def norm(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    return str(value).strip()


def read_report(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    headers = tuple(norm(x) for x in (rows[0] if rows else ()))
    data = [tuple(norm(x) for x in row) for row in rows[1:]]
    return headers, data


def latest_two(folder):
    files = []
    for path in folder.glob("*.xlsx"):
        match = DATE_RE.search(path.name)
        if match:
            files.append((match.group(1), path))
    return [p for _, p in sorted(files, key=lambda item: item[0])[-2:]]


def main():
    summary = []
    details = []
    for shop, biz_dir in TARGETS:
        folder = ROOT / "output" / shop / biz_dir
        files = latest_two(folder)
        if len(files) < 2:
            summary.append([shop, biz_dir, "缺少两版文件", "", "", 0, 0, 0, ""])
            continue
        old_path, new_path = files
        old_headers, old_rows = read_report(old_path)
        new_headers, new_rows = read_report(new_path)
        old_set, new_set = set(old_rows), set(new_rows)
        added, removed = new_set - old_set, old_set - new_set
        header_status = "一致" if old_headers == new_headers else "字段有差异"
        summary.append([
            shop, biz_dir, old_path.name, new_path.name,
            header_status, len(old_rows), len(new_rows), len(added), len(removed),
            "无数据变化" if not added and not removed else "存在数据差异",
        ])
        for change, rows in (("新增", added), ("旧版独有", removed)):
            for row in sorted(rows):
                details.append([shop, biz_dir, old_path.name, new_path.name, change, *row])

    out_dir = ROOT / "output" / "报表差异"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "报表差异报告_2026-08-27.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "差异汇总"
    ws.append(["店铺", "业务", "旧版文件", "新版文件", "字段", "旧版行数", "新版行数", "新增行", "旧版独有行", "结论"])
    for row in summary:
        ws.append(row)
    detail_ws = wb.create_sheet("行级差异")
    max_cols = max((len(row) for row in details), default=5)
    detail_ws.append(["店铺", "业务", "旧版文件", "新版文件", "变化类型"] + [f"字段{i}" for i in range(1, max_cols - 4)])
    for row in details:
        detail_ws.append(row)
    wb.save(out_path)
    print(out_path)
    for row in summary:
        print(" | ".join(map(str, row)))


if __name__ == "__main__":
    main()
