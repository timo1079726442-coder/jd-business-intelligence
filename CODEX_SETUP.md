# Codex 迁移交接说明（2026-08-27）

本项目为京东商智/京准通/京麦数据导出系统。本文档供 codex 环境接手使用。

## 1. 安装依赖

```bash
pip install -r requirements.txt
```

核心依赖：pandas 3.0.5 / openpyxl 3.1.5 / requests / numpy / xlrd / xlsxwriter / msoffcrypto-tool 6.0.0

## 2. 目录结构

```
京东数据导出-2/
├── main.py                  # 业务类（商智7/京准通6/京麦2）+ 注册表
├── run_daily.py             # ★每日调度器（config 驱动，增量补缺）
├── config_download_reader.py# ★业务下载配置加载器
├── excel_master.py          # 总表 collect/flush/rebuild
├── db_utils.py              # SQLite 入库 save_to_db
├── biz_config_loader.py     # 业务/店铺配置加载
├── auth_loader.py           # Cookie/h5st 鉴权加载（JSON 优先）
├── config/
│   ├── config.xlsx          # ★「业务下载配置」sheet 驱动日期+策略
│   └── {SHOP}_*_cookie.json # 鉴权文件（敏感，不入库）
├── output/
│   ├── 总表/{pin}_商智/京准通/京麦总表.xlsx   # 唯一交付物
│   └── {店}/{业务}/{业务}_{日期}.xlsx        # 单表（京准通/京麦仍落盘）
└── data/jd_report.db        # SQLite 业务库
```

## 3. 核心架构（新增机制）

### 3.1 每日调度器 run_daily.py
- 读 config.xlsx「业务下载配置」sheet → 逐报表按策略跑 → DB → 总表
- **策略**：
  - `单日`：扫描缺失日期（最新日期+1 ~ 今天-1）逐日补（京准通快车/商智）
  - `近3天/近7天/近30天`：滚动窗口覆盖（京麦、京准通全站营销）
  - `月度`：整段覆盖（月初手工）
- **收尾 flush_shop 增量写总表**（不 rebuild，快 200 倍）

### 3.2 总表机制（excel_master.py）
- 总表 = DB 的 Excel 视图，3 文件：商智/京准通/京麦
- **表头保持京东原始列**（无 stat_date/shop_pin/report_date 元数据列）
- flush_shop 增量写（按 stat_date_alias 删旧 + append）

### 3.3 业务类（main.py）
- 商智 5 类：不落盘单表，df → save_to_db → 总表
- 京准通/京麦：仍落盘 xlsx，run_daily 按文件内日期列拆分入库

## 4. 运行

```bash
# 干跑看计划（不真调业务）
set SHOP_PIN=miyo-周
set SHOP_ID=MIYO箱包旗舰店
set AUTH_LOADER=1
python run_daily.py --shop MIYO箱包旗舰店 --dry-run

# 真跑（补缺 + 增量写总表）
python run_daily.py --shop MIYO箱包旗舰店

# 全店跑
python run_daily.py
```

退出码：0 成功 / 1 部分失败 / 2 鉴权过期 / 3 系统错误

## 5. 配置「业务下载配置」sheet

| 列 | 说明 |
|---|---|
| 模块 | 京准通 / 京麦 / 商智 |
| 报表 | 展示名 |
| 业务key | 中文全名（与 main.py 注册表一致） |
| 起始日期/结束日期 | 月度/单日策略用；Excel 日期单元格（datetime） |
| 启用 | 下拉：是/否 |
| 覆盖策略 | 下拉：单日/月度/近3天/近7天/近30天 |

## 6. 鉴权（敏感）

- Cookie/h5st 都在 `config/{SHOP}_*_cookie.json` 和 `{SHOP}_*_h5st.json`
- 格式：CDP 抓包日志（AuthLoader 自动抽取）
- h5st 30 分钟过期；京麦订单/售后每次跑需新鲜 h5st
- **绝不提交鉴权文件到 git**（.gitignore 已覆盖）

## 7. 业务能力边界

- **京准通快车（快车推广/快车订单）**：支持单日查询（单日补缺）
- **京准通全站营销（计划/效果/全店计划/全店效果）**：不支持单日（单日返回空），必须区间覆盖（近30天/月度）
- **京准通全站数据 T+N 延迟**：报表截止到昨天前几天，增量模式逐日自动补
- **京麦**：近30天覆盖（订单状态会更新）；订单明细需短信密码（IMAP 监听，超时询问用户）
- **商智**：单日查询，按缺失日期补

## 8. 需人工注意

1. 京麦订单 h5st 过期 → 需影刀 CDP 重抓
2. 京麦订单导出 10 分钟限流（两次任务间隔）
3. 总表京准通 44 万行手动清理超时 → 用 flush 触发自动清理
4. IMAP 密码监听超时 → 先问用户是否已收到短信
