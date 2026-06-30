# software-copyright-design-spec

> 从任意代码项目生成软著申请三件套：申请表 + 源代码 + 设计说明书。

项目地址：https://github.com/antianpeng/software-copyright-design-spec

---

## 解决什么问题

申请软著需要三份材料：著作权申请表、源代码材料（前30页+后30页）、软件设计说明书。整理这些材料耗时且容易出错——申请表26个字段要填对，代码要按规则截取，设计说明书要有流程图。

这个 Skill 让 Claude Code 读你的项目代码，自动生成这三份文档。

## 产出物

| 文件 | 格式 | 内容 |
|------|------|------|
| `{软件全称}.doc` | .doc | 著作权申请表（26行表格） |
| `{软件全称}--源代码.docx` | .docx | 合并源代码（前30页+后30页） |
| `{软件全称}--软件设计说明书.docx` | .docx | 7章设计说明 + Mermaid/Graphviz 流程图 |

## 安装

### 1. 克隆仓库

```bash
git clone https://github.com/antianpeng/software-copyright-design-spec.git
```

### 2. 安装依赖

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli   # mmdc --version
winget install Graphviz.Graphviz          # dot -V
```

### 3. 加载到 Claude Code

```bash
# 方式一：在仓库目录启动
claude --plugin-dir .

# 方式二：指定路径
claude --plugin-dir /path/to/software-copyright-design-spec
```

## 使用

在 Claude Code 中打开你的项目，然后说：

```
使用 software-copyright-design-spec 生成当前项目的软著申请材料
```

Claude 会按 6 步引导你完成：

### Step 1: 精读代码

Claude 自动分析你的项目，提取技术栈、API 端点、核心逻辑、数据模型、表结构等。

### Step 2: 生成构建脚本

将 `build_template.py` 复制到输出目录，重命名为 `build_all.py`。

### Step 3: 填写配置

只需改脚本中 CONFIG 区的 TODO 项：

```python
SOFTWARE_NAME = "XXX管理系统软件"    # 必须以"软件/系统/平台"结尾
VERSION = "V1.0"                    # VX.X 或 VX.X.X
PROJECT_ROOT = r"你的项目路径"
OUTPUT_DIR = r"输出目录"
```

其余字段（著作权人、开发环境、运行环境等）Claude 会帮你填，你确认即可。

### Step 4: 定制图表和设计说明书

Claude 从代码逻辑推导流程图和函数表，生成 Mermaid 图表。你确认或修改。

### Step 5: 运行

```bash
python build_all.py
```

### Step 6: 校验

- [ ] 申请表 26 行完整，无上家公司信息
- [ ] 源代码 60 页，页眉含软件名+版本号
- [ ] 设计说明书 7 章，嵌入 ≥8 张图
- [ ] 三文档软件名和版本号一致
- [ ] 所有图表从真实代码推导

## 字段规则

| 字段 | 要求 |
|------|------|
| 软件名称 | 以"软件/系统/平台"结尾 |
| 版本号 | `VX.X` 或 `VX.X.X` |
| 开发完成日期 | 公司成立后 ≥3 个月 |
| 源程序量 | 全部源码行数（含空行） |
| 主要功能 | 500-1300 字 |
| 技术特点 | ≤100 字 |

## 源代码选取原则

选 Controller/Service/Model + 配置文件 + 前端路由，总行数 ≥3000。代码必须来自真实项目，不编造。

## 常见问题

| 问题 | 解法 |
|------|------|
| Graphviz 中文乱码 | subgraph label 用英文，node label 可以中文 |
| .doc 被 Word 锁住 | 关 Word 重跑，脚本会报 PermissionError |
| 申请表有上家公司信息 | 脚本先清空所有行再填值 |
| 设计说明书不放图 | 必须嵌入 PNG |
| 流程图编造 | 从代码逻辑推导，不凭空捏造 |

## 文件结构

```
software-copyright-design-spec/
├── SKILL.md              # Skill 定义和工作流
├── build_template.py     # 通用构建模板
└── README.md             # 本文件
```

## 开源协议

MIT License
