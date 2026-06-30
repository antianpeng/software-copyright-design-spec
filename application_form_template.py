# -*- coding: utf-8 -*-
"""
Reusable application form (.doc) generator for software copyright.
Copy this file to your project directory and fill in the values.

Usage:
    python application_form_template.py
"""
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


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


def generate_application_form(software_name, fields, output_path):
    """
    Generate application form .doc following the standard template.

    Args:
        software_name: Full software name (软件全称)
        fields: list of (label, value) tuples, 26 items
        output_path: output .doc file path
    """
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # Title
    title = doc.add_paragraph()
    title.alignment = 1  # CENTER
    title.space_after = Pt(12)
    run = title.add_run("公司软件著作权信息登记采集表")
    set_run_font(run, "宋体", 16, bold=True)

    # Table
    table = doc.add_table(rows=len(fields), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, (label, value) in enumerate(fields):
        cell_label = table.cell(i, 0)
        cell_value = table.cell(i, 1)
        cell_label.width = Cm(7)
        cell_value.width = Cm(10)

        for p in cell_label.paragraphs:
            p.clear()
        r0 = cell_label.paragraphs[0].add_run(label)
        set_run_font(r0)

        for p in cell_value.paragraphs:
            p.clear()
        r1 = cell_value.paragraphs[0].add_run(value)
        set_run_font(r1)

        set_cell_shading(cell_label, "D9E2F3")

    doc.save(output_path)
    print(f"[OK] {output_path}")


# ─── Standard 26 fields template ─────────────────────────────────────

STANDARD_FIELDS = [
    # (field_label, default_value)
    # 1. Fill in actual company/personal info
    ("著作权人姓名或名称（必填）", "待填写"),
    ("公司成立日期（必填）", "待填写"),
    ("营业执照注册号、身份证号或事业单位法人证书编号（必填）", "待填写"),
    ("公司地址（必填）", "待填写"),
    # 5. Software name
    ("软件名称（必填）\n1、软件名称必须以“软件或系统或平台”结尾\n2、可参考：“公司简称+产品用途和功能+软件、系统或平台”形式进行命名", "待填写"),
    # 6. Version
    ("版本号（必填）\n按照VX.X或VX.X.X形式填写", "V1.0"),
    # 7. Upgrade
    ("是否是升级版本（必填）", "否"),
    # 8. Short name
    ("软件简称（选填，若有简称，可填写）", ""),
    # 9. Category
    ("软件分类（必填）只能选一个", "应用软件"),
    # 10. Completion date
    ("开发完成日期（必填）\n选定的日期需在公司成立日期之后推迟至少三个月以上", "待填写"),
    # 11. Published
    ("发表状态（必填）", "未发表"),
    # 12-13. First publish
    ("首次发表时间（必填，若尚未发表，不填）", ""),
    ("首次发表地点（必填，若尚未发表，不填）", "城市："),
    # 14-15. Hardware
    ("开发的硬件环境（必填）", "待填写"),
    ("运行的硬件环境（必填）", "待填写"),
    # 16. Dev OS
    ("开发该软件的操作系统（必填）", "待填写"),
    # 17. Dev tools
    ("软件开发环境/开发工具（必填）", "待填写"),
    # 18. Run platform
    ("该软件的运行平台/操作系统（必填）", "待填写"),
    # 19. Runtime
    ("软件运行支撑环境/支持软件（必填）", "待填写"),
    # 20. Language
    ("编程语言（必填）\n与提供的源代码对应", "待填写"),
    # 21. Lines
    ("源程序量（必填）\n指软件源程序量的总行数（单位：行）", "待填写"),
    # 22. Purpose
    ("开发目的（50字以内）（必填）", "待填写"),
    # 23. Domain
    ("面向领域/行业（必填）", "待填写"),
    # 24. Main functions (500-1300 chars)
    ("主要功能（必填）注意不要过于简单，需详细介绍主要功能（限500个字）", "待填写"),
    # 25. Tech features (<=100 chars)
    ("技术特点（必填）（限100个字）", "待填写"),
    # 26. Tech tags
    ("软件技术特点（必填）（最多选3项）", "应用软件"),
]


if __name__ == "__main__":
    # Example usage - customize before running
    name = "XXX系统V1.0"
    fields = list(STANDARD_FIELDS)  # Copy and modify
    # fields[4] = ("软件名称...", "实际软件名称")
    # fields[5] = ("版本号...", "V1.0")
    # ... fill in all fields

    output = os.path.join(".", f"{name}.doc")
    generate_application_form(name, fields, output)
