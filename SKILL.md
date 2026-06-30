---
name: software-copyright-design-spec
description: Use when generating software copyright (软著) application materials that require a design specification (软件设计说明书) with architecture diagrams, flowcharts, and module design. Triggers on 软著设计说明书, 软著带图, 软著流程图, or software copyright with diagrams.
---

# 软著设计说明书生成

从代码项目自动生成三份软著申请材料：申请表(.doc)、合并源代码(.docx)、软件设计说明书(.docx)。

## 核心原则

**图表是设计说明书的灵魂。** 没有流程图、架构图的设计说明书不符合软著审核要求。所有图表必须从真实代码逻辑推导，不能用文字描述代替。

## 前置依赖

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli   # mmdc
winget install Graphviz.Graphviz          # dot
```

验证：`mmdc --version` 和 `dot -V`

## 工作流

### Step 1: 精读代码

从项目中提取关键信息，形成业务理解：

| 来源 | 提取 |
|------|------|
| `pom.xml` / `package.json` | 技术栈、版本号 |
| Controller 文件 | API 端点、模块划分 |
| Service 文件 | 核心业务逻辑、算法 |
| Domain/Model 文件 | 数据实体、字段定义 |
| SQL 文件 | 数据库表结构 |
| 路由文件 | 前端页面结构 |
| 通信模块 (Netty/WebSocket) | 协议、时序 |

### Step 2: 读模板

读取 `模板/` 目录下的 .doc 模板，理解申请表的表格结构（26行2列）。后续通过复制模板+替换值来生成，确保格式一致。

```python
from docx import Document
doc = Document("模板/xxx.doc")
for i, row in enumerate(doc.tables[0].rows):
    print(f"Row {i}: {row.cells[0].text} -> {row.cells[1].text}")
```

### Step 3: 生成图表

按清单生成 10 张 PNG 图：

| # | 图名 | 工具 | 内容 |
|---|------|------|------|
| 01 | system_architecture | Graphviz | 分层架构图 |
| 02 | module_decomposition | Graphviz | 模块分解图 |
| 03 | login_flow | Mermaid | 登录流程 |
| 04 | hall_management | Mermaid | 营业厅管理流程 |
| 05 | terminal_management | Mermaid | 终端设备管理流程 |
| 06 | resource_publish | Mermaid | 资源下发流程 |
| 07 | queue_management | Mermaid | 排队叫号流程 |
| 08 | netty_sequence | Mermaid | 通信时序图 |
| 09 | database_er | Mermaid | 数据库 ER 图 |
| 10 | data_flow | Mermaid | 数据流图 |

**渲染命令：**
```bash
mmdc -i diagram.mmd -o diagram.png -b white -w 1200
dot -Tpng -Gdpi=150 diagram.dot -o diagram.png
```

### Step 4: 生成三个文档

运行 `build_all.py`（配置在文件顶部），自动产出：

- `{软件全称}.doc` — 申请表
- `{软件全称}--源代码.docx` — 合并源代码（前30+后30页）
- `{软件全称}--软件设计说明书.docx` — 含图表的设计说明书

### Step 5: 校验

- [ ] 申请表 .doc 存在且表格26行完整
- [ ] 源代码 .docx 为60页，页眉含软件名+版本号
- [ ] 设计说明书包含7个一级章节
- [ ] 嵌入 ≥8 张图（架构、模块、流程、时序、ER、数据流）
- [ ] 每个模块有功能说明、流程图、函数表、算法说明
- [ ] 三个文档的软件名称和版本号一致

## 关键避坑

| 坑 | 解法 |
|----|------|
| Graphviz 中文乱码 | subgraph label 必须用英文，node label 可以用中文 |
| 申请表格式错 | 复制模板 .doc 再替换值，不要从零建表 |
| 源代码分成两个文件 | 合并为一个 .docx，取前30+后30页 |
| 设计说明书不放图 | 必须嵌入 PNG，不能用文字描述 |
| 流程图编造 | 必须从真实代码逻辑推导 |
| 函数表不写所属类 | 每个函数标明 Controller/Service 类名 |
| 页眉缺失 | 合并源代码页眉格式：`软件名称 版本号` |
