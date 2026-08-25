# -*- coding: utf-8 -*-
"""imap_config_loader.py
============================================================
QQ 邮箱 IMAP 配置加载器（2026-08-21 项目23 改造）

目的：
    - 兼容旧的 config/imap_config.ini（**保留** - 鉴权敏感信息不入仓）
    - 在 config.xlsx「全局配置」sheet 同步登记元信息：
        - IMAP 授权码后 4 位掩码（**只存后4位**，不存完整授权码，避免泄露）
        - 授权码过期提醒日期（auth_code_expire_remind_date）
        - 授权码备注（auth_code_note，如"2026-08-21 用户授权"）
    - 启动时校验：ini 与 xlsx 后 4 位是否一致；不一致打印强警告
    - 过期日期临近（< 14 天）时打印强警告
    - 过期时（< 7 天或已过期）抛 RuntimeError 阻塞业务

数据源：
    主：config/imap_config.ini（运行时真实鉴权，已被 .gitignore 排除）
    辅：config.xlsx「全局配置」sheet「IMAP」组（提醒/元信息，可入仓）

config.xlsx 推荐结构（写在「全局配置」sheet）：
    分组    | key                          | value
    全局    | imap_host                    | imap.qq.com
    全局    | imap_user                    | 1079726442@qq.com
    IMAP    | auth_code_last4              | ihfj           (授权码后4位)
    IMAP    | auth_code_expire_date        | 2026-11-21     (建议过期日，3个月后)
    IMAP    | auth_code_note               | 2026-08-21 用户授权
    IMAP    | imap_auth_code_path          | config/imap_config.ini  (默认)

读取优先级：
    1. 先读 ini 的 auth_code（**真实值**，运行时必须）
    2. 再读 xlsx 的 IMAP 组做一致性校验 + 过期提醒
    3. xlsx 没有 IMAP 组 → 首次启动会打印提示并自动生成（mask 写入）

使用示例：
    from imap_config_loader import load_imap_config, check_imap_health

    cfg = load_imap_config()            # 一次性返回完整配置 dict
    check_imap_health(cfg)              # 检查过期提醒（启动时调用一次）

    # 在京麦订单业务类内
    cfg = load_imap_config()
    auth_code = cfg["auth_code"]
"""
import os
import re
import configparser
import logging
from datetime import datetime, timedelta
from typing import Optional

import openpyxl

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_INI_PATH = os.path.join(PROJECT_ROOT, "config", "imap_config.ini")
CONFIG_XLSX_PATH = os.path.join(PROJECT_ROOT, "config", "config.xlsx")

# 默认 ini 路径（被本模块查找顺序：参数 > 环境变量 > 默认）
DEFAULT_INI_PATH = os.environ.get("IMAP_INI_PATH", CONFIG_INI_PATH)

# 授权码建议有效期：90 天（QQ 邮箱官方建议每 3-6 个月更新）
DEFAULT_VALID_DAYS = 90

# 到期前提醒阈值
WARN_DAYS = 14   # 14 天内开始警告
ERROR_DAYS = 7    # 7 天内或已过期抛异常


# ====================================================================
#  工具函数
# ====================================================================

def _read_xlsx_imap_meta() -> dict:
    """从 config.xlsx「全局配置」sheet 读取 IMAP 元信息。

    返回 dict，键名：
        - auth_code_last4
        - auth_code_expire_date
        - auth_code_note
        - imap_auth_code_path（默认 config/imap_config.ini）
        - imap_host（兜底）
        - imap_user（兜底）

    缺失返回 {}。
    """
    if not os.path.exists(CONFIG_XLSX_PATH):
        return {}
    try:
        wb = openpyxl.load_workbook(CONFIG_XLSX_PATH, read_only=True, data_only=True)
        # 兼容多种 sheet 名
        ws = None
        for name in ("全局配置", "全局", "config"):
            if name in wb.sheetnames:
                ws = wb[name]
                break
        if ws is None:
            wb.close()
            return {}
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            group = str(row[0]).strip() if row[0] is not None else ""
            key = str(row[1]).strip() if row[1] is not None else ""
            val = row[2]
            if not (group and key and val is not None):
                continue
            if group == "IMAP" or group == "全局":
                result[key] = str(val).strip()
        wb.close()
        return result
    except Exception as e:
        logging.warning(f"⚠️ config.xlsx IMAP 元信息读取失败：{e}")
        return {}


def _parse_date(s: str) -> Optional[datetime]:
    """解析日期字符串（支持 yyyy-mm-dd / yyyy/m/d / yyyymmdd）"""
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# ====================================================================
#  对外 API
# ====================================================================

def load_imap_config(ini_path: str = None) -> dict:
    """加载 IMAP 配置（ini 为主，xlsx 为辅）。

    参数:
        ini_path:  ini 文件路径；None → 用 DEFAULT_INI_PATH

    返回:
        dict 包含：
            - host, port, user, auth_code (来自 ini)
            - use_ssl, folder, sender_filter, subject_keyword, max_wait_seconds, poll_interval_seconds
            - auth_code_last4 (来自 xlsx 或从 ini 推导)
            - auth_code_expire_date (来自 xlsx)
            - auth_code_note (来自 xlsx)
            - config_ini_path
            - source: "ini" 标识真实值来源

    异常:
        FileNotFoundError - ini 文件不存在
        ValueError        - 必填字段缺失
    """
    ini_path = ini_path or DEFAULT_INI_PATH
    if not os.path.isfile(ini_path):
        raise FileNotFoundError(
            f"❌ IMAP 配置文件不存在：{ini_path}\n"
            f"   解决：复制 config/imap_config.ini.example 为 imap_config.ini 后填入授权码"
        )

    cfg = configparser.ConfigParser()
    cfg.read(ini_path, encoding="utf-8")
    if "imap" not in cfg:
        raise ValueError(f"❌ IMAP 配置文件缺少 [imap] section：{ini_path}")

    section = cfg["imap"]
    required = ["host", "port", "user", "auth_code"]
    missing = [k for k in required if not section.get(k)]
    if missing:
        raise ValueError(f"❌ IMAP 配置文件缺失必填字段：{missing}")

    auth_code = section["auth_code"].strip()
    # xlsx 元信息
    meta = _read_xlsx_imap_meta()

    return {
        # 真实鉴权（来自 ini）
        "host": section["host"].strip(),
        "port": int(section["port"]),
        "user": section["user"].strip(),
        "auth_code": auth_code,
        "use_ssl": section.get("use_ssl", "true").strip().lower() in ("1", "true", "yes"),
        "folder": section.get("folder", "INBOX").strip(),
        "sender_filter": section.get("sender_filter", "").strip(),
        "subject_keyword": section.get("subject_keyword", "解压密码").strip(),
        "max_wait_seconds": int(section.get("max_wait_seconds", "300")),
        "poll_interval_seconds": int(section.get("poll_interval_seconds", "5")),
        "config_ini_path": ini_path,
        # 元信息（来自 xlsx）
        "auth_code_last4": meta.get("auth_code_last4", auth_code[-4:]),
        "auth_code_expire_date": meta.get("auth_code_expire_date", ""),
        "auth_code_note": meta.get("auth_code_note", ""),
    }


def check_imap_health(cfg: dict) -> None:
    """启动时调用：检查 IMAP 配置健康度 + 打印过期提醒。

    行为:
        - 授权码与 xlsx 后 4 位不匹配 → 强警告（不抛异常）
        - xlsx 未登记 expire_date → 提示用户补登记
        - 距离过期 < 14 天 → 警告（WARN）
        - 距离过期 < 7 天 或 已过期 → 抛 RuntimeError

    参数:
        cfg: load_imap_config() 返回的字典
    """
    auth_code = cfg["auth_code"]
    last4_xlsx = cfg.get("auth_code_last4", "")
    last4_actual = auth_code[-4:] if len(auth_code) >= 4 else auth_code
    expire_str = cfg.get("auth_code_expire_date", "")
    note = cfg.get("auth_code_note", "")

    print(f"\n{'='*70}")
    print(f"📬 [IMAP 配置健康检查] {cfg['config_ini_path']}")
    print(f"{'='*70}")
    print(f"  邮箱账号：{cfg['user']}")
    print(f"  授权码后4位：****{last4_actual}（xlsx 登记：{'****' + last4_xlsx if last4_xlsx else '未登记'}）")
    if note:
        print(f"  备注：{note}")

    # 1) 后 4 位一致性校验
    if last4_xlsx and last4_xlsx != last4_actual:
        print(
            f"\n  ⚠️  警告：xlsx 登记的后4位 ({last4_xlsx}) 与 ini 实际值 ({last4_actual}) 不一致！\n"
            f"      建议同步更新 config.xlsx「全局配置」sheet 的 IMAP 组：\n"
            f"        auth_code_last4 = {last4_actual}\n"
            f"        auth_code_note = <填入变更时间，如 '2026-08-21 用户授权'>"
        )

    # 2) 过期日期校验
    if not expire_str:
        print(
            f"\n  💡 提醒：config.xlsx「全局配置」sheet 的 IMAP 组尚未登记 auth_code_expire_date\n"
            f"      建议格式：auth_code_expire_date = {datetime.now().strftime('%Y-%m-%d')}\n"
            f"      （QQ 邮箱授权码建议 3 个月更新一次）\n"
            f"      默认有效期：{DEFAULT_VALID_DAYS} 天"
        )
        print(f"{'='*70}\n")
        return

    expire_dt = _parse_date(expire_str)
    if not expire_dt:
        print(
            f"\n  ⚠️  警告：xlsx auth_code_expire_date 格式不合法（{expire_str}）\n"
            f"      推荐格式：YYYY-MM-DD（如 2026-11-21）"
        )
        print(f"{'='*70}\n")
        return

    today = datetime.now()
    days_left = (expire_dt - today).days

    if days_left < 0:
        # 已过期：抛异常阻塞
        print(
            f"\n  ❌ 严重警告：QQ 邮箱 IMAP 授权码已过期 {-days_left} 天！\n"
            f"      过期日：{expire_str}\n"
            f"      \n"
            f"      请立即更新：\n"
            f"        1. 登录 QQ 邮箱网页版 → 设置 → 账户 → 开启 IMAP/SMTP → 重新生成授权码\n"
            f"        2. 修改 config/imap_config.ini 的 [imap].auth_code 字段\n"
            f"        3. 同步更新 config.xlsx「全局配置」sheet 的 IMAP 组：\n"
            f"           auth_code_last4 = <新授权码后4位>\n"
            f"           auth_code_expire_date = <新过期日（建议今天 +90 天）>\n"
            f"           auth_code_note = <如 '2026-08-21 用户授权'>"
        )
        print(f"{'='*70}\n")
        raise RuntimeError(
            f"❌ QQ 邮箱 IMAP 授权码已过期 {-days_left} 天（{expire_str}），"
            f"京麦订单明细将无法解压。请按上述步骤更新后重试。"
        )

    elif days_left <= ERROR_DAYS:
        # 即将过期（< 7 天）
        print(
            f"\n  ❌ 警告：QQ 邮箱 IMAP 授权码将在 {days_left} 天后过期（{expire_str}）\n"
            f"      建议尽快更新（QQ 邮箱授权码通常 3-6 个月需重置一次）\n"
            f"      更新步骤：见上方'已过期'提示"
        )
        print(f"{'='*70}\n")
        raise RuntimeError(
            f"❌ QQ 邮箱 IMAP 授权码将在 {days_left} 天后过期，请先更新再跑京麦订单业务。"
        )

    elif days_left <= WARN_DAYS:
        # 临近过期（< 14 天）
        print(
            f"\n  ⚠️  提醒：QQ 邮箱 IMAP 授权码将在 {days_left} 天后过期（{expire_str}）\n"
            f"      建议提前更新，避免京麦订单明细任务中断"
        )
        print(f"{'='*70}\n")

    else:
        # 正常
        print(f"  ✅ 授权码有效，剩余 {days_left} 天（到期 {expire_str}）")
        print(f"{'='*70}\n")


# ====================================================================
#  CLI（调试用）
# ====================================================================

if __name__ == "__main__":
    import sys
    try:
        cfg = load_imap_config()
        check_imap_health(cfg)
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"[ERR] {e}")
        sys.exit(1)