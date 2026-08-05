# 📂 文件夹命名说明：python-code-gen

## 基本信息

| 项 | 内容 |
|---|------|
| **文件夹名** | `python-code-gen` |
| **命名含义** | "Python" + "代码（Code）" + "生成（Gen=Generate）" |
| **完整含义** | Python代码生成Skill |
| **所属** | `.trae/skills/` |
| **创建原因** | 生成面向编程小白的Python代码，强制完整中文注释、配置区与逻辑区分离 |

## 用途

存放Python代码生成规范：
- 代码固定结构（import块→业务配置区→逻辑代码）
- 注释要求（英文术语+中文释义）
- 安全约束（禁止硬编码Cookie/签名/token）

## 文件清单

| 文件 | 作用 |
|------|------|
| `SKILL.md` | Python代码生成Skill主文件 |
| `FOLDER_NAME.md` | 本文件（命名说明）|

## 命名规范

✅ 命名规范：`<语言>-<产物>-<动作>`，缩写友好
- `python-code-gen` ← 完整形式
- `python-codegen` ← 也可以（去掉连字符）