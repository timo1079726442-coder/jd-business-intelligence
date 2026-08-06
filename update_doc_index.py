# -*- coding: utf-8 -*-
"""
update_doc_index.py
============================================================
项目文档索引自动更新工具。

按全局规则：
- 扫描项目内全部 .md / .py 文档
- 跳过：以 _ 开头的临时文件、.git、__pycache__、cdp_user_data、辅助脚本（test_*.py / check_*.py / find_*.py / analyze_*.py / git_*.py / create_*.py 等临时脚本）
- 重复处理：相对路径已存在则跳过，不重复入库
- 类别自动推断：根据文件名 / 路径匹配已知类别（如 AGENTS.md → 全局规则、SKILL.md → Skill）
- 用途简述：自动从文件头 1-3 行提取摘要

使用方式：
    # 手动全量同步
    python update_doc_index.py

    # 后台异步调用（不阻塞主流程）
    from update_doc_index import async_update_index
    async_update_index()  # 启动 daemon 线程，异步扫描
"""

# ============================================================
# 一、import 导入区
# ============================================================
import os
import re
import sys
import threading
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# 二、【业务可修改配置区】
# ============================================================

PROJECT_ROOT = Path(r"D:\CODE\trae\traespace\FYA箱包旗舰店")
OUTPUT_XLSX = PROJECT_ROOT / "docs" / "项目文档索引.xlsx"

# 扫描目标后缀
SCAN_EXTENSIONS = {".md", ".py"}

# 完全跳过的目录（不递归扫描）
SKIP_DIRS = {
    ".git", "__pycache__", "cdp_user_data", "node_modules",
    "venv", ".venv", "dist", "build", ".idea", ".vscode",
}

# 跳过的具体文件（按文件名）
SKIP_FILES = {
    "cdp_network_log.json",  # 已在 .gitignore
}

# 跳过的文件名前缀（临时文件）
SKIP_PREFIXES = ("_",)

# 跳过的 glob 模式（按 .gitignore 临时文件规则）
# 这些是项目内辅助脚本，不算"项目文档"
SKIP_GLOBS = (
    "test_*.py", "check_*.py", "find_*.py",
    "analyze_*.py", "git_*.py", "create_*.py",
)
# 例外：create_config_xlsx.py 是项目工具脚本，需要入库；另外 create_doc_index_xlsx.py
# 生成第一个索引，也是项目工具，保留
SKIP_GLOBB_EXCEPTIONS = {
    "create_config_xlsx.py",
    "create_doc_index_xlsx.py",
    "update_doc_index.py",
}

# 类别推断规则（按精确文件名）
CATEGORY_RULES = {
    # 全局规则
    "AGENTS.md": "全局规则",
    "agents.md": "全局规则",
    "agents_old_backup.md": "全局规则",
    # 架构规范
    "FOLDERS.md": "架构规范",
    "工作区命名与架构规范.md": "架构规范",
    # 业务知识
    "全局复利的踩坑日志.md": "业务知识",
    "API 实现逻辑说明.md": "业务知识",
    # FOLDER_NAME.md（文件夹说明）
    "FOLDER_NAME.md": "文件夹说明",
}

# 路径包含的类别推断
PATH_CATEGORY_RULES = [
    (".trae/skills/", "Skill"),
    (".trae/", "Skill配置"),
]


# ============================================================
# 三、底层固定逻辑区
# ============================================================

def _should_skip_dir(dir_path: Path) -> bool:
    """判断目录是否跳过。"""
    return dir_path.name in SKIP_DIRS


def _should_skip_file(file_path: Path) -> bool:
    """判断文件是否跳过。"""
    name = file_path.name

    # 1. 跳过的具体文件
    if name in SKIP_FILES:
        return True

    # 2. 跳过的前缀（_xxx.py 临时文件）
    if any(name.startswith(p) for p in SKIP_PREFIXES):
        return True

    # 3. 跳过 glob 模式（test_*.py / check_*.py / create_*.py 等）
    if name in SKIP_GLOBB_EXCEPTIONS:
        return False
    for glob_pattern in SKIP_GLOBS:
        # 简单 glob 匹配：前缀 + 扩展名
        # "*.py" 是 4 个字符，不是 3 个
        if glob_pattern.endswith("*.py"):
            prefix = glob_pattern[:-4]  # 去尾 "*.py" 保留前缀
            if name.endswith(".py") and name.startswith(prefix):
                return True

    # 4. 仅收 .md / .py
    if file_path.suffix.lower() not in SCAN_EXTENSIONS:
        return True

    return False


def _extract_usage(file_path: Path) -> str:
    """从文件头 1-3 行提取用途简述。"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            lines = []
            for _ in range(10):
                line = f.readline()
                if not line:
                    break
                line = line.strip()
                # 跳过 docstring 边界、空行
                if not line or line in ('"""', "'''"):
                    continue
                lines.append(line)
                if len(lines) >= 3:
                    break
    except Exception:
        return "（读取文件头失败）"

    if not lines:
        return "（空文件）"

    # 简单拼接前 3 行
    summary = " ".join(lines)
    # 截断到 200 字
    if len(summary) > 200:
        summary = summary[:200] + "..."
    return summary


def _infer_category(file_path: Path, rel_path: str) -> str:
    """根据文件名/路径推断类别。"""
    name = file_path.name

    # 1. 精确文件名匹配
    if name in CATEGORY_RULES:
        return CATEGORY_RULES[name]

    # 2. 路径包含匹配
    for keyword, category in PATH_CATEGORY_RULES:
        if keyword in rel_path:
            return category

    # 3. SKILL.md 一定是 Skill
    if name == "SKILL.md":
        return "Skill"

    # 4. FOLDER_NAME.md 一定是文件夹说明
    if name == "FOLDER_NAME.md":
        return "文件夹说明"

    # 5. 文档类 .md 文件
    if file_path.suffix == ".md":
        return "文档"

    # 6. Python 文件分类
    if file_path.suffix == ".py":
        if name == "main.py":
            return "主程序"
        if name == "jd_cdp_capture.py":
            return "脚本"
        if name.startswith("create_") and name.endswith(".py"):
            return "工具脚本"
        return "脚本"

    return "未分类"


def _scan_project():
    """扫描项目，生成 (相对路径, 绝对路径, 文件名, 类别, 用途简述) 列表。"""
    entries = []
    for root, dirs, files in os.walk(PROJECT_ROOT):
        # 过滤要递归的目录
        dirs[:] = [d for d in dirs if not _should_skip_dir(Path(d))]

        root_path = Path(root)
        for filename in files:
            file_path = root_path / filename
            if _should_skip_file(file_path):
                continue
            rel_path = str(file_path.relative_to(PROJECT_ROOT)).replace("\\", "/")
            category = _infer_category(file_path, rel_path)
            usage = _extract_usage(file_path)
            entries.append((rel_path, str(file_path), filename, category, usage))

    # 按相对路径排序
    entries.sort(key=lambda x: x[0])
    return entries


def _load_existing_xlsx():
    """读取现有 xlsx，返回 (headers, existing_rel_paths) 集合。"""
    if not OUTPUT_XLSX.exists():
        return None, set()

    try:
        wb = load_workbook(OUTPUT_XLSX)
        if "�� 文档总索引" not in wb.sheetnames:
            return None, set()
        ws = wb["�� 文档总索引"]
        # 第一行是表头，第 4 列是相对路径
        existing_paths = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 4 and row[3]:
                existing_paths.add(row[3])
        return wb, existing_paths
    except Exception as e:
        print(f"[WARN] 读取现有 xlsx 失败：{e}")
        return None, set()


def _build_styles():
    """构造 xlsx 样式。"""
    header_font = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_font = Font(name="微软雅黑", size=10)
    cell_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    return {
        "header_font": header_font,
        "header_fill": header_fill,
        "header_align": header_align,
        "cell_font": cell_font,
        "cell_align": cell_align,
        "border": border,
    }


# 类别颜色映射
CATEGORY_COLORS = {
    "全局规则": "FFE699",
    "架构规范": "C6E0B4",
    "业务知识": "BDD7EE",
    "Skill": "F4B084",
    "Skill配置": "F4B084",
    "文件夹说明": "D9E1F2",
    "文档": "FFF2CC",
    "主程序": "C6E0B4",
    "脚本": "E2EFDA",
    "工具脚本": "D5A6BD",
    "未分类": "FFFFFF",
}


def _apply_styles_to_ws(ws, styles, start_row, end_row, category_col_idx=3):
    """对一行数据应用样式。"""
    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(1, 8):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = styles["cell_font"]
            cell.alignment = styles["cell_align"]
            cell.border = styles["border"]

        # 类别列着色
        category_cell = ws.cell(row=row_idx, column=category_col_idx)
        category = category_cell.value
        fill_color = CATEGORY_COLORS.get(category, "FFFFFF")
        category_cell.fill = PatternFill(
            start_color=fill_color, end_color=fill_color, fill_type="solid"
        )


def _write_header(ws, styles):
    """写入表头。"""
    headers = ["序号", "文件名", "类别", "相对路径", "绝对路径", "用途简述", "状态/备注"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = styles["header_font"]
        cell.fill = styles["header_fill"]
        cell.alignment = styles["header_align"]
        cell.border = styles["border"]


def _build_category_summary_sheet(wb, categories):
    """构建分类汇总 sheet。"""
    # 删除旧的分类汇总
    if "�� 分类汇总" in wb.sheetnames:
        del wb["�� 分类汇总"]

    ws = wb.create_sheet("�� 分类汇总", 1)

    headers = ["类别", "文件数量", "说明"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color="BFBFBF"),
            right=Side(style="thin", color="BFBFBF"),
            top=Side(style="thin", color="BFBFBF"),
            bottom=Side(style="thin", color="BFBFBF"),
        )

    descriptions = {
        "全局规则": "Trae Agent 全局底线，禁止写业务细节；每次启动自动加载。",
        "架构规范": "项目目录结构、命名规则、代码组织方式。",
        "业务知识": "踩坑日志、API 实现逻辑文档等业务沉淀。",
        "Skill": "Trae Skill 业务分组（按需加载，不被全局自动注入）。",
        "Skill配置": "Trae Skill 相关配置目录文档。",
        "文件夹说明": "每个目录的 FOLDER_NAME.md，描述目录用途。",
        "文档": "其他通用 .md 文档。",
        "主程序": "项目主入口（如 main.py）。",
        "脚本": "Python 业务脚本（如 Playwright 抓包脚本）。",
        "工具脚本": "项目辅助工具（create_*.py 等）。",
        "未分类": "未匹配到具体类别的文档。",
    }

    for row_idx, (cat, count) in enumerate(sorted(categories.items()), start=2):
        ws.cell(row=row_idx, column=1, value=cat).font = Font(name="微软雅黑", size=10)
        ws.cell(row=row_idx, column=2, value=count).font = Font(name="微软雅黑", size=10)
        ws.cell(row=row_idx, column=3, value=descriptions.get(cat, "")).font = Font(name="微软雅黑", size=10)

        for col_idx in range(1, 4):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = Border(
                left=Side(style="thin", color="BFBFBF"),
                right=Side(style="thin", color="BFBFBF"),
                top=Side(style="thin", color="BFBFBF"),
                bottom=Side(style="thin", color="BFBFBF"),
            )

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 60
    ws.row_dimensions[1].height = 28


def _set_column_widths(ws):
    """设置列宽。"""
    column_widths = {
        "A": 6,
        "B": 38,
        "C": 12,
        "D": 50,
        "E": 80,
        "F": 60,
        "G": 38,
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.row_dimensions[1].height = 28


def _set_row_heights(ws, end_row):
    """设置行高。"""
    for row_idx in range(2, end_row + 1):
        ws.row_dimensions[row_idx].height = 38


def update_index(verbose=True):
    """核心函数：扫描项目目录，对比现有 xlsx，追加新文件。

    入参：
        verbose: 是否打印日志
    返回：
        (新增条数, 跳过条数)
    """
    if verbose:
        print(f"[INFO] 开始扫描项目目录...")
    scanned = _scan_project()
    if verbose:
        print(f"[INFO] 扫描到 {len(scanned)} 个有效文档")

    # 读取现有 xlsx
    wb, existing_paths = _load_existing_xlsx()
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    if wb is None:
        # xlsx 不存在或损坏，新建
        wb = Workbook()
        wb.remove(wb.active)

    # 找到/创建「�� 文档总索引」sheet
    if "�� 文档总索引" in wb.sheetnames:
        ws = wb["�� 文档总索引"]
    else:
        ws = wb.create_sheet("�� 文档总索引", 0)
        styles = _build_styles()
        _write_header(ws, styles)
        ws.freeze_panes = "A2"

    # 找当前最大序号
    max_seq = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            try:
                n = int(str(row[0]))
                if n > max_seq:
                    max_seq = n
            except (ValueError, TypeError):
                pass

    # 追加新文件
    added = 0
    skipped = 0
    new_categories = {}
    styles = _build_styles()

    for rel_path, abs_path, filename, category, usage in scanned:
        if rel_path in existing_paths:
            skipped += 1
            continue

        max_seq += 1
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1, value=max_seq)
        ws.cell(row=new_row, column=2, value=filename)
        ws.cell(row=new_row, column=3, value=category)
        ws.cell(row=new_row, column=4, value=rel_path)
        ws.cell(row=new_row, column=5, value=abs_path)
        ws.cell(row=new_row, column=6, value=usage)
        ws.cell(row=new_row, column=7, value="自动入库 @ " + datetime.now().strftime("%Y-%m-%d"))

        # 应用样式
        _apply_styles_to_ws(ws, styles, new_row, new_row)

        added += 1
        new_categories[category] = new_categories.get(category, 0) + 1
        if verbose:
            print(f"  + {rel_path} → {category}")

    # 统计全部分类
    all_categories = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[2]:
            all_categories[row[2]] = all_categories.get(row[2], 0) + 1

    # 更新列宽和行高
    _set_column_widths(ws)
    _set_row_heights(ws, ws.max_row)

    # 重建分类汇总
    _build_category_summary_sheet(wb, all_categories)

    # 保存
    wb.save(OUTPUT_XLSX)

    if verbose:
        print(f"\n[OK] 索引已更新：{OUTPUT_XLSX}")
        print(f"     新增 {added} 条，跳过 {skipped} 条")
        if new_categories:
            print(f"     新增分类：{new_categories}")

    return added, skipped


def async_update_index():
    """后台异步触发索引更新（不阻塞 main 启动）。"""
    def _worker():
        try:
            update_index(verbose=False)
        except Exception as e:
            print(f"[WARN] 后台索引更新失败：{e}")

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    return thread


def main():
    """入口函数。"""
    print("=" * 60)
    print("项目文档索引 - 自动更新工具")
    print("=" * 60)
    print(f"\n项目根目录：{PROJECT_ROOT}")
    print(f"索引输出：  {OUTPUT_XLSX}")
    print(f"扫描范围：  递归全部 .md / .py")
    print(f"跳过规则：  {SKIP_DIRS} / {SKIP_PREFIXES} / {SKIP_GLOBS}")
    print()

    added, skipped = update_index(verbose=True)

    print(f"\n[DONE] 新增 {added} 条，跳过 {skipped} 条")