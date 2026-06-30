# -*- coding: utf-8 -*-
"""
Reusable diagram generation templates for software copyright materials.
Copy this file to your project directory and customize the content.

Usage:
    python diagram_templates.py

Dependencies:
    pip install python-docx
    npm install -g @mermaid-js/mermaid-cli
    winget install Graphviz.Graphviz
"""
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')

# ─── Configuration ───────────────────────────────────────────────────
# Set these before running
DIAGRAMS_DIR = "./diagrams"
MMDC = "mmdc"  # or full path: r"C:\Users\{user}\AppData\Roaming\npm\mmdc.cmd"
DOT = "dot"    # or full path: r"C:\Program Files\Graphviz\bin\dot.exe"


def run_mermaid(name, mmd_content):
    """Render Mermaid diagram to PNG."""
    os.makedirs(DIAGRAMS_DIR, exist_ok=True)
    mmd_path = os.path.join(DIAGRAMS_DIR, f"{name}.mmd")
    png_path = os.path.join(DIAGRAMS_DIR, f"{name}.png")
    with open(mmd_path, "w", encoding="utf-8") as f:
        f.write(mmd_content)
    result = subprocess.run(
        [MMDC, "-i", mmd_path, "-o", png_path, "-b", "white", "-w", "1200"],
        capture_output=True, text=True, timeout=30
    )
    if result.returncode == 0 and os.path.exists(png_path):
        print(f"  [OK] {name}.png")
    else:
        print(f"  [FAIL] {name}: {result.stderr[:200]}")
    return png_path


def run_graphviz(name, dot_content):
    """Render Graphviz diagram to PNG."""
    os.makedirs(DIAGRAMS_DIR, exist_ok=True)
    dot_path = os.path.join(DIAGRAMS_DIR, f"{name}.dot")
    png_path = os.path.join(DIAGRAMS_DIR, f"{name}.png")
    with open(dot_path, "w", encoding="utf-8") as f:
        f.write(dot_content)
    result = subprocess.run(
        [DOT, "-Tpng", "-Gdpi=150", dot_path, f"-o{png_path}"],
        capture_output=True, text=True, timeout=30
    )
    if result.returncode == 0 and os.path.exists(png_path):
        print(f"  [OK] {name}.png")
    else:
        print(f"  [FAIL] {name}: {result.stderr[:200]}")
    return png_path


# ─── Template: System Architecture (Graphviz) ───────────────────────
def template_architecture(software_name, layers):
    """
    Generate a layered system architecture diagram.

    Args:
        software_name: Display name
        layers: list of (label, nodes) where nodes is list of (id, label)
                e.g. [("客户端层", [("browser", "浏览器")]),
                       ("应用服务层", [("backend", "Spring Boot")])]
    """
    dot = f'digraph architecture {{\n'
    dot += '    rankdir=TB;\n'
    dot += '    node [shape=box, style="rounded,filled", fillcolor="#E8F4FD", fontname="Microsoft YaHei", fontsize=11];\n'
    dot += '    edge [fontname="Microsoft YaHei", fontsize=9];\n\n'

    colors = ["#F5F5F5", "#FFF3E0", "#E8F5E9", "#E3F2FD", "#FCE4EC", "#F3E5F5"]
    edge_color = ["#999999", "#FF9800", "#4CAF50", "#2196F3", "#E91E63", "#9C27B0"]

    prev_cluster = None
    for i, (label, nodes) in enumerate(layers):
        color = colors[i % len(colors)]
        dot += f'    subgraph cluster_{i} {{\n'
        dot += f'        label="{label}"; style="rounded,dashed"; fillcolor="{color}";\n'
        for nid, nlabel in nodes:
            dot += f'        {nid} [label="{nlabel}"];\n'
        dot += '    }\n\n'

    # Connect layers top to bottom
    for i in range(len(layers) - 1):
        _, src_nodes = layers[i]
        _, dst_nodes = layers[i + 1]
        if src_nodes and dst_nodes:
            dot += f'    {src_nodes[0][0]} -> {dst_nodes[0][0]};\n'

    dot += '}\n'
    return run_graphviz("01_system_architecture", dot)


# ─── Template: Module Decomposition (Graphviz) ──────────────────────
def template_modules(modules):
    """
    Generate module decomposition diagram.

    Args:
        modules: list of (cluster_label, fillcolor, bordercolor, submodules)
                 where submodules is list of (id, label, fillcolor)
    """
    dot = 'digraph modules {\n'
    dot += '    rankdir=TB;\n'
    dot += '    node [shape=box, style="rounded,filled", fontname="Microsoft YaHei", fontsize=10];\n\n'

    for i, (label, fill, border, subs) in enumerate(modules):
        dot += f'    subgraph cluster_{i} {{\n'
        dot += f'        label="{label}"; style="rounded"; fillcolor="{fill}"; color="{border}";\n'
        for sid, slabel, sfill in subs:
            dot += f'        {sid} [label="{slabel}" fillcolor="{sfill}"];\n'
        dot += '    }\n\n'

    dot += '}\n'
    return run_graphviz("02_module_decomposition", dot)


# ─── Template: Business Flow (Mermaid) ───────────────────────────────
def template_flow(name, steps):
    """
    Generate a business flow flowchart.

    Args:
        name: filename prefix
        steps: list of Mermaid flow steps, e.g.
               ["A[开始] --> B{判断}", "B -->|是| C[操作]", "B -->|否| D[结束]"]
    """
    mmd = "flowchart TD\n"
    for step in steps:
        mmd += f"    {step}\n"
    return run_mermaid(name, mmd)


# ─── Template: Sequence Diagram (Mermaid) ────────────────────────────
def template_sequence(name, participants, messages):
    """
    Generate a sequence diagram.

    Args:
        name: filename prefix
        participants: list of (id, label)
        messages: list of (from, to, label, note=None)
    """
    mmd = "sequenceDiagram\n"
    for pid, plabel in participants:
        mmd += f"    participant {pid} as {plabel}\n"
    mmd += "\n"
    for msg in messages:
        if len(msg) == 4 and msg[3]:
            mmd += f"    Note over {msg[0]},{msg[1]}: {msg[3]}\n"
        mmd += f"    {msg[0]}->{msg[1]}: {msg[2]}\n"
    return run_mermaid(name, mmd)


# ─── Template: ER Diagram (Mermaid) ──────────────────────────────────
def template_er(name, entities):
    """
    Generate an ER diagram.

    Args:
        name: filename prefix
        entities: list of (entity_name, fields, relationships)
                  where fields is list of (type, name, constraint)
                  and relationships is list of (target, cardinality, label)
    """
    mmd = "erDiagram\n"
    for ename, fields, rels in entities:
        mmd += f"    {ename} {{\n"
        for ftype, fname, fconstraint in fields:
            mmd += f"        {ftype} {fname} {fconstraint}\n"
        mmd += "    }\n\n"
    for ename, _, rels in entities:
        for target, cardinality, label in rels:
            mmd += f"    {ename} {cardinality} {target} : \"{label}\"\n"
    return run_mermaid(name, mmd)


# ─── Example Usage ───────────────────────────────────────────────────

def example_generate_all():
    """Example: generate all diagrams for a typical web application."""
    print("Generating diagrams...")

    # 1. Architecture
    template_architecture("示例系统", [
        ("客户端层", [("browser", "浏览器")]),
        ("代理层", [("nginx", "Nginx")]),
        ("应用层", [("app", "Spring Boot"), ("ws", "WebSocket")]),
        ("数据层", [("mysql", "MySQL"), ("redis", "Redis")]),
    ])

    # 2. Modules
    template_modules([
        ("前端 (Vue.js)", "#E3F2FD", "#1565C0", [
            ("login", "登录模块", "#BBDEFB"),
            ("main", "业务模块", "#BBDEFB"),
            ("api", "API层", "#90CAF9"),
        ]),
        ("后端 (Spring Boot)", "#E8F5E9", "#2E7D32", [
            ("ctrl", "Controller", "#C8E6C9"),
            ("svc", "Service", "#A5D6A7"),
            ("dao", "Mapper", "#81C784"),
        ]),
    ])

    # 3. Login flow
    template_flow("03_login_flow", [
        "A[打开浏览器] --> B[显示登录页]",
        "B --> C[输入账号密码]",
        "C --> D{验证}",
        "D -->|成功| E[跳转首页]",
        "D -->|失败| F[提示错误]",
        "F --> C",
    ])

    print("Done!")


if __name__ == "__main__":
    example_generate_all()
