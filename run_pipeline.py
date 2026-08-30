# -*- coding: utf-8 -*-
"""京东数据统一流水线唯一推荐入口。

使用示例：
    python run_pipeline.py --shop MIYO --dry-run
    python run_pipeline.py --shop MIYO --biz "京麦订单明细_完整一键导出"
    python run_pipeline.py --shop MIYO --start 2026-08-24 --end 2026-08-26

影刀未来只需保存鉴权后调用本文件；采集、补数、SQLite、MySQL、Derived、Metric
都由本入口编排。影刀本阶段仍不修改。
"""

import argparse  # 命令行参数解析库。
import json  # 输出机器可读 JSON Summary。
import sys  # 将稳定退出码返回给影刀或计划任务。
from pathlib import Path  # 安全处理项目路径。

from pipeline.orchestrator import PipelineOrchestrator
from pipeline.status import EXIT_PARAMETER_ERROR


def _configure_utf8_stdio() -> None:
    """统一 Windows 控制台 UTF-8 输出，避免中文 Summary 被错误编码。"""
    for stream in (getattr(sys, "stdout", None), getattr(sys, "stderr", None)):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except (OSError, ValueError):
                pass


def _split_values(values: list[str] | None) -> list[str]:
    """兼容重复 --shop/--biz 与逗号分隔写法。"""
    result = []
    for value in values or []:
        result.extend(item.strip() for item in value.split(",") if item.strip())
    return result


def main() -> int:
    """解析参数、执行流水线，并在 stdout 最后一行打印 JSON Summary。"""
    _configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="京东数据统一流水线（唯一推荐入口）")
    parser.add_argument("--shop", action="append", help="店铺短名或全名，例如 MIYO；可重复或逗号分隔")
    parser.add_argument("--biz", action="append", help="业务 biz_key；可重复或逗号分隔")
    parser.add_argument("--start", help="调试/补数起始日期，格式 YYYY-MM-DD；必须与 --end 一起使用")
    parser.add_argument("--end", help="调试/补数结束日期，格式 YYYY-MM-DD；必须与 --start 一起使用")
    parser.add_argument("--trigger-source", choices=["manual", "rpa", "scheduled"], default="manual", help="触发来源")
    parser.add_argument("--dry-run", action="store_true", help="仅生成计划和鉴权状态；不请求京东 API、不写 SQLite/MySQL")
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent
    try:
        pipeline = PipelineOrchestrator(project_root, trigger_source=args.trigger_source, dry_run=args.dry_run)
        summary = pipeline.run(
            shop_filters=_split_values(args.shop),
            biz_filters=_split_values(args.biz),
            start_override=args.start,
            end_override=args.end,
        )
        # 最后一行固定为 JSON，供影刀/RPA 直接读取，不需解析自然语言日志。
        print(json.dumps(summary, ensure_ascii=False))
        return int(summary["exit_code"])
    except (ValueError, FileNotFoundError) as exc:
        print(json.dumps({"status": "PARAMETER_ERROR", "error": str(exc), "exit_code": EXIT_PARAMETER_ERROR}, ensure_ascii=False))
        return EXIT_PARAMETER_ERROR
    except Exception as exc:
        print(json.dumps({"status": "SYSTEM_ERROR", "error_type": type(exc).__name__, "exit_code": EXIT_PARAMETER_ERROR}, ensure_ascii=False))
        return EXIT_PARAMETER_ERROR


if __name__ == "__main__":
    sys.exit(main())
