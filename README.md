# software-copyright-design-spec

> Claude Code Skill：从任意项目生成软著申请三件套。

项目地址：https://github.com/antianpeng/software-copyright-design-spec

---

## 做什么

读你的项目代码，自动生成三份软著申请材料：

| 文件 | 说明 |
|------|------|
| `{软件全称}.doc` | 著作权申请表（26行表格，从模板填充） |
| `{软件全称}--源代码.docx` | 合并源代码（前30页+后30页，从真实项目抽取） |
| `{软件全称}--软件设计说明书.docx` | 7章设计说明 + 嵌入 Mermaid/Graphviz 流程图 |

## 为什么做

软著申请需要三份材料，整理起来耗时：申请表26个字段要填对，代码要按规则截取，设计说明书要有流程图。这个 Skill 让 Claude Code 帮你做，你只需要确认和微调。

## 安装

```bash
git clone https://github.com/antianpeng/software-copyright-design-spec.git
cd software-copyright-design-spec
```

### 依赖

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli   # 流程图渲染
winget install Graphviz.Graphviz          # 架构图渲染
```

### 加载到 Claude Code

```bash
claude --plugin-dir .
```

## 使用

在你的项目目录中启动 Claude Code，说：

```
使用 software-copyright-design-spec 生成当前项目的软著申请材料
```

Claude 会读你的代码，生成构建脚本 `build_all.py`，你确认配置后运行：

```bash
python build_all.py
```

详细工作流（6步校验清单、字段规则、图表定制）见 [SKILL.md](SKILL.md)。

## 避坑

| 问题 | 解法 |
|------|------|
| Graphviz 中文乱码 | subgraph label 用英文，node label 可以中文 |
| .doc 被 Word 锁住 | 关 Word 重跑 |
| 申请表有上家公司信息 | 脚本先清空所有行再填值 |
| 流程图编造 | 从代码逻辑推导，不凭空捏造 |

## 文件结构

```
├── SKILL.md              # Skill 定义 + 完整工作流
├── build_template.py     # 通用构建模板（复制后填配置运行）
├── LICENSE               # MIT
└── README.md             # 本文件
```

## 开源协议

[MIT License](LICENSE)
