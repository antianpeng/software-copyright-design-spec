---
name: software-copyright-design-spec
description: Use when generating software copyright (软著) application materials for ANY project. Produces application form, merged source code, and design specification with diagrams. Triggers on 软著, 软著申请, 软著设计说明书, software copyright, or 软著材料.
---

# 软著申请材料生成（通用版）

从任意代码项目自动生成三份软著申请材料。适用于任何技术栈（Java/Python/Go/Node等）。

## 产出物

| 文件 | 格式 | 说明 |
|------|------|------|
| `{软件全称}.doc` | .doc | 申请表（26行表格） |
| `{软件全称}--源代码.docx` | .docx | 合并源代码（前30+后30页） |
| `{软件全称}--软件设计说明书.docx` | .docx | 7章设计说明+嵌入图表 |

## 前置依赖

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli   # 验证: mmdc --version
winget install Graphviz.Graphviz          # 验证: dot -V
```

## 完整工作流

### Step 1: 精读代码，形成业务理解

**目的：** 理解这个软件是做什么的、怎么实现的，为后续填写申请表和写设计说明书积累素材。

必读文件清单：

| 读什么 | 提取什么 |
|--------|---------|
| `pom.xml` / `package.json` / `go.mod` | 技术栈、依赖版本、模块划分 |
| Controller / Router 文件 | 所有 API 端点、业务模块划分 |
| Service / UseCase 文件 | 核心业务逻辑、算法、数据处理 |
| Entity / Model / Domain 文件 | 数据实体、字段定义、关系 |
| SQL / Migration 文件 | 数据库表结构 |
| 配置文件 (`application.yml` 等) | 端口、中间件、环境配置 |
| 通信模块 (WebSocket/MQTT/TCP) | 协议、消息格式、时序 |
| 前端路由 (`router/`) | 页面结构、模块划分 |

**输出：** 在脑中形成「这个系统做什么 → 有哪些模块 → 每个模块怎么实现」的完整图景。

### Step 2: 读模板

读取 `模板/` 目录下的 .doc 模板，理解申请表的表格结构（26行2列，字段名+字段值）。

```python
from docx import Document
doc = Document("模板/xxx.doc")
for i, row in enumerate(doc.tables[0].rows):
    print(f"Row {i}: {row.cells[0].text} -> {row.cells[1].text}")
```

### Step 3: 填写 build_template.py 的 CONFIG 区

打开 `build_template.py`，修改 CONFIG 区域的所有 `TODO` 项：

```python
# 必填项
SOFTWARE_NAME = "你的软件全称"      # 必须以"软件/系统/平台"结尾
VERSION = "V1.0"                    # 格式 VX.X 或 VX.X.X
PROJECT_ROOT = r"你的项目路径"
TEMPLATE_DOC = r"模板/xxx.doc"      # 申请表模板路径
OUTPUT_DIR = r"输出目录"

# 申请表26个字段值（参考模板字段说明填写）
FILL_VALUES = { 0: "著作权人", 1: "成立日期", ... }

# 源代码文件列表（从项目中选出核心文件）
SELECTED_FILES = [ "path/to/file1.java", ... ]
```

**字段填写规则：**

| 字段 | 规则 |
|------|------|
| 软件名称 | 必须以"软件/系统/平台"结尾 |
| 版本号 | `VX.X` 或 `VX.X.X` |
| 开发完成日期 | 在公司成立日期之后至少3个月 |
| 源程序量 | 全部源码总行数（含空行），用 `wc -l` 统计 |
| 主要功能 | 500-1300字，详细介绍每个模块 |
| 技术特点 | ≤100字 |
| 开发/运行硬件环境 | ≤50字 |

**源代码选取原则：**
- 选核心业务文件（Controller、Service、Model）
- 包含配置文件（pom.xml、package.json）
- 前端选路由和核心组件
- 总行数足够60页（每页50行 = 3000行以上）

### Step 4: 定制设计说明书内容

`build_template.py` 中的 `generate_design_spec()` 包含7章模板结构。**必须根据你的项目定制以下内容：**

#### 4.1 引言（3个小节）
- 编写目的：一句话说明本文档是XXX的软件设计说明书
- 项目背景：这个软件解决什么问题
- 术语定义：项目特有的技术术语

#### 4.2 总体设计
- 系统架构：描述你的技术架构（B/S、C/S、微服务等）
- 技术栈表格：从 pom.xml/package.json 提取
- 模块划分：按你的项目结构调整

#### 4.3 接口设计
- REST API 表格：从 Controller 文件提取所有端点
- WebSocket / TCP / MQ 接口：按实际通信方式填写

#### 4.4 模块设计（核心，每个模块4个小节）
对每个主要业务模块：
1. 功能说明 — 这个模块做什么
2. 处理流程图 — 用 Mermaid 画业务流程
3. 关键函数表 — 函数名 | 所属类 | 功能说明
4. 算法说明 — 核心业务算法

#### 4.5 运行设计
- 运行流程
- 通信时序图
- 数据流图

#### 4.6 出错处理设计
- 接口异常、通信异常、数据库异常

#### 4.7 数据库设计
- ER 关系图
- 核心表结构表格

### Step 5: 运行生成

```bash
python build_template.py
```

脚本按顺序执行：
1. 生成10张图表 → `diagrams/` 目录
2. 生成申请表 → `.doc`
3. 合并源代码 → `.docx`
4. 生成设计说明书 → `.docx`

### Step 6: 校验

- [ ] 申请表 .doc 存在且表格26行完整
- [ ] 源代码 .docx 为60页，页眉含软件名+版本号
- [ ] 设计说明书包含7个一级章节
- [ ] 嵌入 ≥8 张图（架构、模块、流程、时序、ER、数据流）
- [ ] 每个模块有功能说明、流程图、函数表、算法说明
- [ ] 三个文档的软件名称和版本号一致
- [ ] 所有图表从真实代码推导，非 AI 编造

## 关键避坑

| 坑 | 解法 |
|----|------|
| Graphviz 中文乱码 | subgraph label 必须用英文，node label 可以用中文 |
| 申请表格式错 | 复制模板 .doc 再替换值，不要从零建表 |
| 源代码分成两个文件 | 合并为一个 .docx，取前30+后30页 |
| 设计说明书不放图 | 必须嵌入 PNG，不能用文字描述代替 |
| 流程图编造 | 必须从真实代码逻辑推导 |
| 函数表不写所属类 | 每个函数标明 Controller/Service 类名 |
| 页眉缺失 | 合并源代码页眉格式：`软件名称 版本号` |
| 主要功能太短 | 必须500-1300字，逐个模块详细介绍 |
| 版本号格式错 | 必须 `VX.X` 或 `VX.X.X` |
| 开发日期早于公司成立 | 开发完成日期必须在公司成立日期之后3个月以上 |

## 文件结构

```
software-copyright-design-spec/
  SKILL.md                  # 本文件（方法论）
  build_template.py         # 通用生成模板（改配置就能跑）
```
