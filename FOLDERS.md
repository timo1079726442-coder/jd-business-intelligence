# 📂 项目目录结构与命名说明（总索引）

> 每个文件夹内都有自己的 `FOLDER_NAME.md` 详细说明。本文件是总索引，便于快速了解项目结构。

## 项目根目录：`FYA箱包旗舰店`

| 文件/目录 | 命名含义 | 作用 |
|----------|----------|------|
| `FYA箱包旗舰店/` | 京东店铺名 | 项目根目录（店铺：FYA8888） |
| `.trae/` | Trae配置目录 | Trae IDE识别的项目级配置 |
| `agents.md` | 全局规则文件 | 全局铁律精简版，禁止堆业务 |
| `main.py` | 主程序入口 | 所有业务代码集中在一个文件 |
| `config/` | 配置文件 | 存放Cookie、Excel配置 |
| `docs/` | 文档目录 | 项目技术文档、API清单 |
| `jd_api/` | 京东API包 | 已废弃，遗留目录（已合并到main.py） |
| `logs/` | 日志目录 | 程序运行日志（按日期） |
| `output/` | 输出目录 | 导出的Excel文件 |

## Skill目录结构：`.trae/`

```
.trae/
├── FOLDER_NAME.md        ← 命名说明（点开头 = Trae自动识别）
└── skills/
    ├── FOLDER_NAME.md
    ├── jd-api-analyze/   ← 京东API分析（商智/京麦/京准通）
    │   ├── FOLDER_NAME.md
    │   └── SKILL.md
    ├── git-safe-operate/ ← Git安全操作
    │   ├── FOLDER_NAME.md
    │   └── SKILL.md
    └── python-code-gen/  ← Python代码生成
        ├── FOLDER_NAME.md
        └── SKILL.md
```

## 命名规范总览

| 类型 | 规范 | 示例 |
|------|------|------|
| **业务领域** | 全小写、连字符分隔 | `jd-api-analyze` |
| **通用目录** | 业界标准名 | `config/`、`docs/`、`logs/`、`output/` |
| **Python包** | 下划线分隔 | `jd_api/`（已废弃） |
| **隐藏目录** | 点开头 | `.trae/`、`.git/`、`.gitignore` |
| **Skill子目录** | `<业务>-<功能>` | `git-safe-operate` |

## 📝 FOLDER_NAME.md 说明

每个目录里都有一个 `FOLDER_NAME.md` 文件，包含：
- 文件夹名 / 命名含义 / 完整含义
- 用途和文件清单
- 命名规范说明
- ⚠️ 必要的警告（如jd_api/已废弃）

这个约定的好处：
- ✅ 打开任何目录都能立刻看到命名原因
- ✅ 不会污染 agents.md
- ✅ 易于维护，新人快速上手