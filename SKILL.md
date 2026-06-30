---
name: software-copyright-design-spec
description: Use when generating software copyright (软著) application materials for ANY project. Produces application form, merged source code, and design specification with diagrams. Triggers on 软著, 软著申请, 软著设计说明书, software copyright, or 软著材料.
---

# 软著申请材料生成

从任意代码项目自动生成三份软著申请材料。适用于任何技术栈。

## 产出物

| 文件 | 格式 | 说明 |
|------|------|------|
| `{软件全称}.doc` | .doc | 申请表（26行表格） |
| `{软件全称}--源代码.docx` | .docx | 合并源代码（前30+后30页） |
| `{软件全称}--软件设计说明书.docx` | .docx | 7章设计说明+嵌入图表 |

## 依赖

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli   # mmdc --version
winget install Graphviz.Graphviz          # dot -V
```

## 工作流（6步）

### Step 1: 精读代码

从项目中提取业务理解：

| 读什么 | 提取什么 |
|--------|---------|
| `pom.xml` / `package.json` | 技术栈、版本 |
| Controller / Router | API 端点、模块划分 |
| Service | 核心逻辑、算法 |
| Entity / Model | 数据实体、字段 |
| SQL / Migration | 表结构 |
| 配置文件 | 端口、中间件 |
| 前端路由 | 页面结构 |

### Step 2: 复制 build_template.py

复制到输出目录，重命名为 `build_all.py`。

### Step 3: 填 CONFIG 区

只需改 CONFIG 区的 TODO 项：

```python
SOFTWARE_NAME = "你的软件全称"      # 必须以"软件/系统/平台"结尾
VERSION = "V1.0"                    # VX.X 或 VX.X.X
PROJECT_ROOT = r"你的项目路径"
OUTPUT_DIR = r"输出目录"
FILL_VALUES = { ... }               # 26个字段，前4项留空
SELECTED_FILES = [ ... ]            # 核心源代码文件
```

**字段规则：**

| 字段 | 要求 |
|------|------|
| 软件名称 | 以"软件/系统/平台"结尾 |
| 版本号 | `VX.X` 或 `VX.X.X` |
| 开发完成日期 | 公司成立后≥3个月 |
| 源程序量 | 全部源码行数（含空行） |
| 主要功能 | 500-1300字 |
| 技术特点 | ≤100字 |

**源代码选取：** 选 Controller/Service/Model + 配置文件 + 前端路由，总行数≥3000。

### Step 4: 定制图表和设计说明书

修改 `generate_diagrams()` 中的图表内容和 `generate_design_spec()` 中的文档内容。

每个模块需要4样东西：功能说明 + 流程图 + 函数表 + 算法说明。

### Step 5: 运行

```bash
python build_all.py
```

### Step 6: 校验

- [ ] 申请表26行完整，无上家公司信息
- [ ] 源代码60页，页眉含软件名+版本号
- [ ] 设计说明书7章，嵌入≥8张图
- [ ] 三文档软件名和版本号一致
- [ ] 所有图表从真实代码推导

## 避坑

| 坑 | 解法 |
|----|------|
| Graphviz 中文乱码 | subgraph label 用英文，node label 可以中文 |
| .doc 被 Word 锁住 | 脚本会报 PermissionError，关 Word 重跑 |
| 申请表有上家公司信息 | 脚本先清空所有行再填值 |
| 设计说明书不放图 | 必须嵌入 PNG |
| 流程图编造 | 从代码逻辑推导 |
| 函数表不写所属类 | 每个函数标明类名 |
| 主要功能太短 | 500-1300字，逐模块介绍 |

## 文件结构

```
software-copyright-design-spec/
  SKILL.md              # 方法论
  build_template.py     # 通用模板
```
