# -*- coding: utf-8 -*-
"""
软著申请材料一键生成脚本（通用模板）
生成：申请表(.doc) + 合并源代码(.docx) + 软件设计说明书(.docx)

用法：
    1. 修改下方 CONFIG 区域
    2. 修改 generate_design_spec() 中的项目内容
    3. python build_template.py
"""
import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ═══════════════════════════════════════════════════════════════════════
# CONFIG — 只需改这里
# ═══════════════════════════════════════════════════════════════════════

# ── 必填 ──
SOFTWARE_NAME = "XXX管理系统软件"    # 必须以"软件/系统/平台"结尾
VERSION = "V1.0"                      # VX.X 或 VX.X.X
PROJECT_ROOT = r"D:\code\your-project" # 项目代码根目录
OUTPUT_DIR = r"D:\2026软著\你的输出目录"

# ── 申请表26个字段 ──
# 空字符串 = 留空让用户手填（前4项通常留空）
# FILL_VALUES[9] 开发完成日期：必须在公司成立日期之后3个月以上
# FILL_VALUES[20] 源程序量：全部源码总行数（含空行）
# FILL_VALUES[23] 主要功能：500-1300字
# FILL_VALUES[24] 技术特点：≤100字
FILL_VALUES = {
    0: "",    # 著作权人（用户填）
    1: "",    # 公司成立日期（用户填）
    2: "",    # 营业执照注册号（用户填）
    3: "",    # 公司地址（用户填）
    4: SOFTWARE_NAME,
    5: VERSION,
    6: "否",
    7: "",    # 软件简称（选填）
    8: "应用软件",
    9: "",    # TODO: 开发完成日期
    10: "未发表",
    11: "",
    12: "城市：",
    13: "",   # TODO: 开发硬件环境（≤50字）
    14: "",   # TODO: 运行硬件环境（≤50字）
    15: "",   # TODO: 开发操作系统
    16: "",   # TODO: 开发环境/工具
    17: "",   # TODO: 运行平台/操作系统
    18: "",   # TODO: 运行支撑环境
    19: "",   # TODO: 编程语言
    20: "",   # TODO: 源程序量（行）
    21: "",   # TODO: 开发目的（50字以内）
    22: "",   # TODO: 面向领域/行业
    23: "",   # TODO: 主要功能（500-1300字）
    24: "",   # TODO: 技术特点（≤100字）
    25: "应用软件",
}

# ── 源代码文件列表（相对于 PROJECT_ROOT）──
# 选核心业务文件，总行数 ≥ 3000（60页×50行）
SELECTED_FILES = [
    # TODO: 列出你的源代码文件
    # "pom.xml",
    # "src/main/java/com/example/controller/UserController.java",
    # "src/main/java/com/example/service/UserService.java",
    # "src/main/java/com/example/entity/User.java",
    # "src/router/index.js",
]

# ── 工具路径（一般不用改）──
MMDC = "mmdc"  # 或 r"C:\Users\{user}\AppData\Roaming\npm\mmdc.cmd"
DOT = "dot"    # 或 r"C:\Program Files\Graphviz\bin\dot.exe"
TEMPLATE_DOC = r""  # TODO: 填写申请表 .doc 模板的绝对路径
DIAGRAMS_DIR = os.path.join(OUTPUT_DIR, "diagrams")

# ═══════════════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════════════

def set_run_font(run, name="宋体", size=9, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), name)

def set_cell_shading(cell, color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_page_header(section, text):
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.clear()
    run = p.add_run(text)
    set_run_font(run, "宋体", 9)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def add_heading_styled(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "黑体"
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), "黑体")
    return h

def add_para(doc, text, font_name="宋体", font_size=12, bold=False, indent=0, space_after=6):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, font_name, font_size, bold)
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(space_after)
    return p

def add_table_from_rows(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(h)
        set_run_font(run, "宋体", 10, bold=True)
        set_cell_shading(cell, "D9E2F3")
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(val))
            set_run_font(run, "宋体", 10)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    return table

def run_mermaid(name, mmd_content):
    os.makedirs(DIAGRAMS_DIR, exist_ok=True)
    mmd_path = os.path.join(DIAGRAMS_DIR, f"{name}.mmd")
    png_path = os.path.join(DIAGRAMS_DIR, f"{name}.png")
    with open(mmd_path, "w", encoding="utf-8") as f:
        f.write(mmd_content)
    result = subprocess.run([MMDC, "-i", mmd_path, "-o", png_path, "-b", "white", "-w", "1200"],
                            capture_output=True, text=True, timeout=60)
    if result.returncode == 0 and os.path.exists(png_path):
        print(f"  [OK] {name}.png")
    else:
        print(f"  [FAIL] {name}: {result.stderr[:200]}")
    return png_path

def run_graphviz(name, dot_content):
    os.makedirs(DIAGRAMS_DIR, exist_ok=True)
    dot_path = os.path.join(DIAGRAMS_DIR, f"{name}.dot")
    png_path = os.path.join(DIAGRAMS_DIR, f"{name}.png")
    with open(dot_path, "w", encoding="utf-8") as f:
        f.write(dot_content)
    result = subprocess.run([DOT, "-Tpng", "-Gdpi=150", dot_path, f"-o{png_path}"],
                            capture_output=True, text=True, timeout=30)
    if result.returncode == 0 and os.path.exists(png_path):
        print(f"  [OK] {name}.png")
    else:
        print(f"  [FAIL] {name}: {result.stderr[:200]}")
    return png_path

def _add_image(doc, diagram_name):
    img_path = os.path.join(DIAGRAMS_DIR, f"{diagram_name}.png")
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=Inches(5.5))
    else:
        add_para(doc, f"[图表缺失: {diagram_name}.png]", "宋体", 10)

# ═══════════════════════════════════════════════════════════════════════
# 图表生成
# ═══════════════════════════════════════════════════════════════════════

def generate_diagrams():
    """
    生成全部10张图表。
    TODO: 根据你的项目修改图表内容。Graphviz subgraph label 必须用英文！
    """
    print("\n=== 生成图表 ===")

    # 01 系统架构图 (Graphviz) — TODO: 改层次和节点
    run_graphviz("01_system_architecture", '''digraph architecture {
    rankdir=TB;
    node [shape=box, style="rounded,filled", fillcolor="#E8F4FD", fontname="Microsoft YaHei", fontsize=11];

    subgraph cluster_client { label="Client Layer"; style="rounded,dashed"; fillcolor="#F5F5F5";
        browser [label="浏览器"];
    }
    subgraph cluster_app { label="Application Layer"; style="rounded,dashed"; fillcolor="#E8F5E9";
        backend [label="应用服务器"];
    }
    subgraph cluster_data { label="Data Layer"; style="rounded,dashed"; fillcolor="#FCE4EC";
        db [label="数据库"];
    }
    browser -> backend -> db;
}''')

    # 02 模块分解图 (Graphviz) — TODO: 改模块名
    run_graphviz("02_module_decomposition", '''digraph modules {
    rankdir=TB;
    node [shape=box, style="rounded,filled", fontname="Microsoft YaHei", fontsize=10];
    subgraph cluster_m1 { label="Module Group 1"; style="rounded"; fillcolor="#E3F2FD";
        a [label="Module A" fillcolor="#BBDEFB"];
    }
    subgraph cluster_m2 { label="Module Group 2"; style="rounded"; fillcolor="#E8F5E9";
        b [label="Module B" fillcolor="#C8E6C9"];
    }
    a -> b;
}''')

    # 03 登录流程 — TODO: 改登录逻辑
    run_mermaid("03_login_flow", """flowchart TD
    A[打开浏览器] --> B[输入账号密码]
    B --> C{验证}
    C -->|失败| D[提示错误]
    D --> B
    C -->|成功| E[生成Token]
    E --> F[跳转首页]""")

    # 04-07 核心业务流程 — TODO: 改为你的4个核心模块
    run_mermaid("04_module1_flow", """flowchart TD
    A[进入模块1] --> B[查看列表]
    B --> C{选择操作}
    C -->|新增| D[填写信息]
    D --> E[保存]
    C -->|编辑| F[修改信息]
    F --> E
    C -->|删除| G[确认删除]""")

    run_mermaid("05_module2_flow", """flowchart TD
    A[进入模块2] --> B[加载数据]
    B --> C{选择操作}
    C -->|操作1| D[执行处理]
    C -->|操作2| E[查询分析]""")

    run_mermaid("06_module3_flow", """flowchart TD
    A[进入模块3] --> B[选择对象]
    B --> C[配置参数]
    C --> D[执行操作]""")

    run_mermaid("07_module4_flow", """flowchart TD
    A[进入模块4] --> B[查看状态]
    B --> C{需要操作?}
    C -->|是| D[执行操作]
    C -->|否| E[返回列表]""")

    # 08 时序图 — TODO: 改参与者和消息
    run_mermaid("08_sequence", """sequenceDiagram
    participant C as 客户端
    participant S as 服务器
    participant D as 数据库
    C->>S: 请求
    S->>D: 查询
    D->>S: 结果
    S->>C: 响应""")

    # 09 ER图 — TODO: 改实体和关系
    run_mermaid("09_database_er", """erDiagram
    USER {
        bigint id PK
        varchar username UK
        varchar password
    }
    ORDER {
        bigint id PK
        bigint user_id FK
        decimal amount
    }
    USER ||--o| ORDER : "creates" """)

    # 10 数据流图 — TODO: 改数据流向
    run_mermaid("10_data_flow", """flowchart LR
    A[输入] --> B[处理]
    B --> C[(存储)]
    C --> D[输出]""")

    print("图表生成完毕")

# ═══════════════════════════════════════════════════════════════════════
# 申请表 (.doc)
# ═══════════════════════════════════════════════════════════════════════

def generate_application_form():
    """复制模板 → 清空所有值 → 填入新值（确保不留上家公司信息）"""
    print("\n=== 生成申请表 ===")
    output_path = os.path.join(OUTPUT_DIR, f"{SOFTWARE_NAME}.doc")

    if not os.path.exists(TEMPLATE_DOC):
        print(f"  [ERROR] 模板不存在: {TEMPLATE_DOC}")
        return

    import shutil
    try:
        shutil.copy2(TEMPLATE_DOC, output_path)
    except PermissionError:
        print(f"  [ERROR] 无法写入 {output_path}，请关闭 Word 后重试")
        return

    doc = Document(output_path)
    table = doc.tables[0]

    # 关键改动：先清空所有行的右列（去掉上家公司信息）
    for i in range(len(table.rows)):
        cell = table.rows[i].cells[1]
        for p in cell.paragraphs:
            for run in p.runs:
                run.text = ""
        if cell.paragraphs:
            cell.paragraphs[0].clear()

    # 再填入新值
    for idx, val in FILL_VALUES.items():
        if val == "" or idx >= len(table.rows):
            continue
        cell = table.rows[idx].cells[1]
        run = cell.paragraphs[0].add_run(val)
        set_run_font(run, "微软雅黑", 14)

    try:
        doc.save(output_path)
        print(f"  [OK] {output_path}")
    except PermissionError:
        print(f"  [ERROR] 无法保存 {output_path}，请关闭 Word 后重试")

# ═══════════════════════════════════════════════════════════════════════
# 合并源代码 (.docx)
# ═══════════════════════════════════════════════════════════════════════

def generate_merged_code():
    """合并源代码为60页（前30+后30），每50行一页"""
    print("\n=== 生成合并源代码 ===")
    output_path = os.path.join(OUTPUT_DIR, f"{SOFTWARE_NAME}--源代码.docx")

    all_lines = []
    for rel_path in SELECTED_FILES:
        full_path = os.path.join(PROJECT_ROOT, rel_path)
        if not os.path.exists(full_path):
            print(f"  [SKIP] {rel_path}")
            continue
        try:
            with open(full_path, "r", encoding="utf-8", errors="ignore") as f:
                lines = [l.rstrip() for l in f if l.strip()]
            all_lines.append(f"// File: {rel_path}")
            all_lines.extend(lines)
            all_lines.append("")
        except Exception as e:
            print(f"  [SKIP] {rel_path}: {e}")

    if not all_lines:
        print("  [ERROR] 没有读取到任何源代码文件")
        return

    PAGE_SIZE = 50
    pages = [all_lines[i:i+PAGE_SIZE] for i in range(0, len(all_lines), PAGE_SIZE)]
    selected_pages = pages[:30] + pages[-30:] if len(pages) > 60 else pages

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    set_page_header(section, f"{SOFTWARE_NAME} {VERSION}")

    for page_idx, page_lines in enumerate(selected_pages):
        if page_idx > 0:
            doc.add_page_break()
        for line in page_lines:
            p = doc.add_paragraph()
            run = p.add_run(line)
            set_run_font(run, "Courier New", 8)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = Pt(10)

    doc.save(output_path)
    print(f"  [OK] {output_path} ({len(selected_pages)} pages, {len(all_lines)} lines)")

# ═══════════════════════════════════════════════════════════════════════
# 软件设计说明书 (.docx)
# ═══════════════════════════════════════════════════════════════════════

def generate_design_spec():
    """7章设计说明书，嵌入图表。TODO: 根据项目填充内容。"""
    print("\n=== 生成软件设计说明书 ===")
    output_path = os.path.join(OUTPUT_DIR, f"{SOFTWARE_NAME}--软件设计说明书.docx")

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # 封面
    for _ in range(6):
        doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(SOFTWARE_NAME)
    set_run_font(run, "黑体", 22, bold=True)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("软件设计说明书")
    set_run_font(run, "黑体", 18)
    ver = doc.add_paragraph()
    ver.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = ver.add_run(f"版本：{VERSION}")
    set_run_font(run, "宋体", 14)
    doc.add_page_break()

    # 目录
    add_heading_styled(doc, "目录", 1)
    for item in ["一、引言", "二、总体设计", "三、接口设计", "四、模块设计",
                  "五、运行设计", "六、出错处理设计", "七、数据库设计"]:
        add_para(doc, item, "宋体", 12)
    doc.add_page_break()

    # 一、引言
    add_heading_styled(doc, "一、引言", 1)
    add_heading_styled(doc, "1.1 编写目的", 2)
    # TODO: 改为你的软件名
    add_para(doc, f"本文档是《{SOFTWARE_NAME}》的软件设计说明书，为软件著作权登记提供技术文档支撑。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "1.2 项目背景", 2)
    # TODO: 描述你的项目背景
    add_para(doc, "TODO: 描述项目背景，这个软件解决什么问题。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "1.3 术语定义", 2)
    # TODO: 列出项目特有术语
    add_para(doc, "TODO: 列出项目特有的技术术语及其定义。", "宋体", 12, indent=0.74)
    doc.add_page_break()

    # 二、总体设计
    add_heading_styled(doc, "二、总体设计", 1)
    add_heading_styled(doc, "2.1 系统架构", 2)
    # TODO: 描述架构
    add_para(doc, "TODO: 描述系统技术架构。", "宋体", 12, indent=0.74)
    _add_image(doc, "01_system_architecture")
    add_heading_styled(doc, "2.2 技术栈", 2)
    # TODO: 填写技术栈
    add_table_from_rows(doc, ["层次", "技术选型"], [
        ("后端", "TODO"), ("前端", "TODO"), ("数据库", "TODO"),
    ], [4, 10])
    add_heading_styled(doc, "2.3 模块划分", 2)
    _add_image(doc, "02_module_decomposition")
    doc.add_page_break()

    # 三、接口设计
    add_heading_styled(doc, "三、接口设计", 1)
    add_heading_styled(doc, "3.1 REST API接口", 2)
    # TODO: 填写API表格
    add_table_from_rows(doc, ["接口路径", "模块", "功能"], [
        ("/api/xxx/**", "TODO", "TODO"),
    ], [4, 3, 6])
    add_heading_styled(doc, "3.2 其他接口", 2)
    add_para(doc, "TODO: WebSocket/TCP/MQ等接口。", "宋体", 12, indent=0.74)
    doc.add_page_break()

    # 四、模块设计（核心）
    add_heading_styled(doc, "四、模块设计", 1)
    # TODO: 为每个核心模块添加4个小节（功能说明+流程图+函数表+算法）
    for mod_num, mod_name in enumerate(["模块1", "模块2", "模块3"], 1):
        add_heading_styled(doc, f"4.{mod_num} {mod_name}", 2)
        add_heading_styled(doc, f"4.{mod_num}.1 功能说明", 3)
        add_para(doc, f"TODO: 描述{mod_name}的功能。", "宋体", 12, indent=0.74)
        add_heading_styled(doc, f"4.{mod_num}.2 处理流程图", 3)
        _add_image(doc, f"0{mod_num+2}_module{mod_num}_flow")
        add_heading_styled(doc, f"4.{mod_num}.3 关键函数", 3)
        add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], [
            ("TODO()", "TODO_Class", "TODO"),
        ], [3, 4, 7])
        add_heading_styled(doc, f"4.{mod_num}.4 算法说明", 3)
        add_para(doc, "TODO: 描述核心算法。", "宋体", 12, indent=0.74)
    doc.add_page_break()

    # 五、运行设计
    add_heading_styled(doc, "五、运行设计", 1)
    add_heading_styled(doc, "5.1 运行流程", 2)
    add_para(doc, "TODO: 描述系统运行流程。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "5.2 通信时序", 2)
    _add_image(doc, "08_sequence")
    add_heading_styled(doc, "5.3 数据流", 2)
    _add_image(doc, "10_data_flow")
    doc.add_page_break()

    # 六、出错处理设计
    add_heading_styled(doc, "六、出错处理设计", 1)
    add_heading_styled(doc, "6.1 接口异常处理", 2)
    add_para(doc, "TODO: 描述接口异常处理。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "6.2 通信异常处理", 2)
    add_para(doc, "TODO: 描述通信异常处理。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "6.3 数据库异常处理", 2)
    add_para(doc, "TODO: 描述数据库异常处理。", "宋体", 12, indent=0.74)
    doc.add_page_break()

    # 七、数据库设计
    add_heading_styled(doc, "七、数据库设计", 1)
    add_heading_styled(doc, "7.1 ER关系图", 2)
    _add_image(doc, "09_database_er")
    add_heading_styled(doc, "7.2 核心表结构", 2)
    # TODO: 填写表结构
    add_table_from_rows(doc, ["表名", "说明", "核心字段"], [
        ("TODO", "TODO", "TODO"),
    ], [3, 3, 8])

    doc.save(output_path)
    print(f"  [OK] {output_path}")

# ═══════════════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════════════

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"软件名称: {SOFTWARE_NAME}")
    print(f"版本号: {VERSION}")
    print(f"输出目录: {OUTPUT_DIR}")

    generate_diagrams()
    generate_application_form()
    generate_merged_code()
    generate_design_spec()

    print("\n=== 全部完成 ===")
    print(f"  1. {SOFTWARE_NAME}.doc")
    print(f"  2. {SOFTWARE_NAME}--源代码.docx")
    print(f"  3. {SOFTWARE_NAME}--软件设计说明书.docx")
    print(f"  4. diagrams/ (图表PNG)")

if __name__ == "__main__":
    main()
