# 📂 文件夹命名说明：git-safe-operate

## 基本信息

| 项 | 内容 |
|---|------|
| **文件夹名** | `git-safe-operate` |
| **命名含义** | "Git" + "安全（Safe）" + "操作（Operate）" |
| **完整含义** | Git命令安全操作校验Skill |
| **所属** | `.trae/skills/` |
| **创建原因** | 防御 `git rm -f` 等高危命令物理删除源码，给AI一个统一的Git安全检查标准 |

## 用途

存放Git命令安全相关规则：
- 高危命令识别（git rm -f、2>/null等）
- 安全等价改写方案（--cached方案）
- shell关键字中文释义
- 小白记忆区分对照表

## 文件清单

| 文件 | 作用 |
|------|------|
| `SKILL.md` | Git安全操作校验Skill主文件 |
| `FOLDER_NAME.md` | 本文件（命名说明）|

## 命名规范

✅ 命名规范：`<工具名>-<核心目的>`
- 工具名在前：`git` / `python` / `shell`
- 核心目的在后：`safe-operate` / `code-gen` / `analyze`