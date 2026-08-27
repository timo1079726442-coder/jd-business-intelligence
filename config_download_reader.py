# -*- coding: utf-8 -*-
"""config_download_reader.py — 业务下载配置加载器（2026-08-26 项目27 新增）

背景：
    用户要求日期设为动态写入配置文件，每个报表下载日期可手工选择，
    京准通按月度覆盖、京麦按 30 天覆盖、商智按单日查缺补日。

功能：
    读 config.xlsx「业务下载配置」sheet → List[DownloadConfig]
    含模块/报表/业务key/起止日期/启用/覆盖策略字段。

覆盖策略：
    - 月度：拉整段（start_date ~ end_date），DB upsert 整月覆盖
    - 30天：拉近 30 天（end=昨天），订单状态最新覆盖
    - 单日：拉配置起止日期（同一天），按日查缺补
"""
import os
import sys
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from typing import List, Optional

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(PROJECT_ROOT, "config", "config.xlsx")

SHEET_NAME = "业务下载配置"


@dataclass
class DownloadConfig:
    """单个报表的下载配置"""
    module: str          # "京准通" / "京麦" / "商智"
    report_name: str     # "快车推广" / "订单明细" / "搜索流量"
    biz_key: str         # 业务 key（中文全名，如「京准通快车自定义报表」）
    start_date: str      # YYYY-MM-DD（月度策略必填；30天策略忽略；单日策略可省略）
    end_date: str        # YYYY-MM-DD（月度策略必填；30天策略忽略；单日策略即补哪一天）
    enabled: bool        # 启用=True / 禁用=False
    strategy: str        # "月度" / "30天" / "单日"

    def resolve_dates(self) -> tuple:
        """根据策略解析实际下载日期范围。

        返回:
            (start_date, end_date) 元组，均为 YYYY-MM-DD 字符串

        策略:
            - 月度: 用配置里的 start_date / end_date（整段）
            - 近3天 / 近7天 / 近30天: end=昨天, start=end-(N-1)（订单状态滚动覆盖）
            - 单日: start_date = end_date = 配置的 end_date（或今天）
        """
        if self.strategy == "月度":
            return self.start_date, self.end_date
        elif self.strategy in ("近3天", "近7天", "近30天"):
            n = {"近3天": 3, "近7天": 7, "近30天": 30}[self.strategy]
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
            start = (datetime.now() - timedelta(days=n)).strftime("%Y-%m-%d")
            return start, yesterday
        elif self.strategy == "单日":
            # 单日策略：start_date=end_date=配置里的日期（默认今天）
            target = self.end_date or self.start_date or datetime.now().strftime("%Y-%m-%d")
            return target, target
        else:
            raise ValueError(f"❌ 未知覆盖策略：{self.strategy}（期望 单日/月度/近3天/近7天/近30天）")


def _str_to_bool(val) -> bool:
    """字符串转 bool（兼容 是/否/True/1/yes/on）"""
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    return s in ("是", "true", "1", "yes", "y", "on")


def _norm_date(v) -> str:
    """Excel 日期单元格（datetime）或字符串 → 'YYYY-MM-DD'。

    背景（2026-08-26）：用户在 Excel 里手工填日期单元格，openpyxl 读出来是
    datetime 对象（如 datetime(2026,8,25)），str() 会变成 '2026-08-25 00:00:00'，
    resolve_dates 用 str 拼接会出错。统一转 'YYYY-MM-DD'。
    """
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    return str(v or "").strip()


def load_download_configs(config_path: Optional[str] = None) -> List[DownloadConfig]:
    """读 config.xlsx「业务下载配置」sheet，返回启用='是' 的 DownloadConfig 列表。

    入参:
        config_path - 可选，默认 config/config.xlsx
    返回:
        List[DownloadConfig]
    异常:
        FileNotFoundError - config 文件不存在
        ValueError - sheet 不存在
    """
    path = config_path or CONFIG_PATH
    if not os.path.exists(path):
        raise FileNotFoundError(f"❌ config 文件不存在：{path}")

    try:
        import openpyxl
    except ImportError:
        raise ImportError("❌ 需要安装 openpyxl：pip install openpyxl")

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"❌ config.xlsx 缺少「{SHEET_NAME}」sheet\n"
            f"   请手工新增 sheet 并填入配置行（列：模块/报表/业务key/起始日期/结束日期/启用/覆盖策略）"
        )

    ws = wb[SHEET_NAME]
    configs = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # 跳过表头
        if not row or not row[0]:
            continue  # 跳过空行
        if len(row) < 7:
            print(f"[WARN] 第 {i+1} 行列数不足（{len(row)}/7），跳过")
            continue
        module, report_name, biz_key, start_date, end_date, enabled, strategy = row[:7]
        cfg = DownloadConfig(
            module=str(module or "").strip(),
            report_name=str(report_name or "").strip(),
            biz_key=str(biz_key or "").strip(),
            start_date=_norm_date(start_date),
            end_date=_norm_date(end_date),
            enabled=_str_to_bool(enabled),
            strategy=str(strategy or "").strip(),
        )
        if not cfg.enabled:
            continue
        if not cfg.biz_key:
            print(f"[WARN] 第 {i+1} 行 biz_key 为空，跳过：{cfg.report_name}")
            continue
        if not cfg.strategy:
            print(f"[WARN] 第 {i+1} 行 strategy 为空，跳过：{cfg.biz_key}")
            continue
        configs.append(cfg)
    wb.close()
    return configs


def list_enabled_by_strategy(strategy: str, config_path: Optional[str] = None) -> List[DownloadConfig]:
    """筛选指定策略的所有启用配置"""
    return [c for c in load_download_configs(config_path) if c.strategy == strategy]


# CLI 调试入口
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="业务下载配置加载器调试")
    parser.add_argument("--strategy", choices=["月度", "30天", "单日"], help="按策略筛选")
    parser.add_argument("--dry-run", action="store_true", help="只读不跑（默认 True）")
    args = parser.parse_args()

    try:
        if args.strategy:
            configs = list_enabled_by_strategy(args.strategy)
        else:
            configs = load_download_configs()
    except (FileNotFoundError, ValueError) as e:
        print(e)
        sys.exit(3)

    print(f"=== 加载到 {len(configs)} 条启用配置 ===")
    for c in configs:
        s, e = c.resolve_dates()
        print(f"  [{c.module}] {c.report_name} ({c.biz_key})")
        print(f"    起始={c.start_date} 结束={c.end_date} 策略={c.strategy} → 实际下载 {s} ~ {e}")
