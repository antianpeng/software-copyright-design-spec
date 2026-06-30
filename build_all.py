# -*- coding: utf-8 -*-
"""
软著申请材料一键生成脚本
生成：申请表(.doc) + 合并源代码(.docx) + 软件设计说明书(.docx)

用法：
    1. 修改下方 CONFIG 区域的配置
    2. 修改 SELECTED_FILES 为你的源代码文件列表
    3. 修改 generate_design_spec() 中的文档内容
    4. python build_all.py
"""
import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ═══════════════════════════════════════════════════════════════════════
# CONFIG — 修改这里
# ═══════════════════════════════════════════════════════════════════════

SOFTWARE_NAME = "基于边缘代理的营业厅运管平台系统"
VERSION = "V3.8.6"
PROJECT_ROOT = r"D:\code\ningxia"
TEMPLATE_DOC = r"D:\2026软著\模板\基于人工智能的电力现场作业音视频合规分析与工单联动系统软件V1.0.doc"
OUTPUT_DIR = r"D:\2026软著\基于边缘代理的营业厅运管平台系统"

# 工具路径（按需修改）
MMDC = r"C:\Users\93452\AppData\Roaming\npm\mmdc.cmd"
DOT = r"C:\Program Files\Graphviz\bin\dot.exe"
DIAGRAMS_DIR = os.path.join(OUTPUT_DIR, "diagrams")

# 申请表字段值（按模板26行顺序，空字符串表示留空让用户填）
# 前4项（著作权人等）通常留空
FILL_VALUES = {
    0: "",   # 著作权人姓名或名称
    1: "",   # 公司成立日期
    2: "",   # 营业执照注册号
    3: "",   # 公司地址
    4: SOFTWARE_NAME,
    5: VERSION,
    6: "否",
    7: "",
    8: "应用软件",
    9: "2026-06-29",
    10: "未发表",
    11: "",
    12: "城市：",
    13: "Intel CPU 16核、内存16GB、硬盘300GB",
    14: "Intel CPU 4核、内存8GB、硬盘100GB",
    15: "Windows 10",
    16: "开发环境: Windows 10/开发工具: Visual Studio Code",
    17: "Windows 10/11 或 macOS 13及以上版本",
    18: "Node.js、npm、现代浏览器",
    19: "Java, JavaScript",
    20: "7195",
    21: "为电力营业厅提供电子设备全生命周期运营管理的信息化手段",
    22: "电力行业 / 营业厅运营管理",
    23: "面向电力行业营业厅的电子设备全生命周期运营管理平台，通过边缘代理网关实现终端设备的远程接入、实时状态监控、媒体资源下发和运维管理，帮助营业厅实现设备管理数字化和运维智能化。",
    24: "基于边缘代理网关的设备接入技术，采用Netty TCP长连接实现设备状态实时采集，支持WebSocket实时通信推送，前后端分离架构，基于Spring Security和JWT的多终端认证体系",
    25: "应用软件",
}

# 源代码文件列表（相对于 PROJECT_ROOT）
SELECTED_FILES = [
    "pom.xml",
    "platform-admin/pom.xml",
    "platform-admin/src/main/java/com/ruoyi/web/controller/system/SysLoginController.java",
    "platform-business/pom.xml",
    "platform-business/src/main/java/com/ruoyi/business/controller/HallInfoController.java",
    "platform-business/src/main/java/com/ruoyi/business/controller/TerminalDeviceController.java",
    "platform-business/src/main/java/com/ruoyi/business/controller/MediaResourceController.java",
    "platform-business/src/main/java/com/ruoyi/business/controller/QueueInfoController.java",
    "platform-business/src/main/java/com/ruoyi/business/controller/DeviceAccessoryController.java",
    "platform-business/src/main/java/com/ruoyi/business/service/IHallInfoService.java",
    "platform-business/src/main/java/com/ruoyi/business/service/impl/HallInfoServiceImpl.java",
    "platform-business/src/main/java/com/ruoyi/business/service/ITerminalDeviceService.java",
    "platform-business/src/main/java/com/ruoyi/business/service/impl/TerminalDeviceServiceImpl.java",
    "platform-business/src/main/java/com/ruoyi/business/service/IMediaResourceService.java",
    "platform-business/src/main/java/com/ruoyi/business/service/impl/MediaResourceServiceImpl.java",
    "platform-business/src/main/java/com/ruoyi/business/service/IQueueInfoService.java",
    "platform-business/src/main/java/com/ruoyi/business/service/impl/QueueInfoServiceImpl.java",
    "platform-business/src/main/java/com/ruoyi/business/domain/HallInfo.java",
    "platform-business/src/main/java/com/ruoyi/business/domain/TerminalDevice.java",
    "platform-business/src/main/java/com/ruoyi/business/domain/MediaResource.java",
    "platform-business/src/main/java/com/ruoyi/business/domain/QueueInfo.java",
    "platform-business/src/main/java/com/ruoyi/business/domain/DeviceAccessory.java",
    "platform-business/src/main/java/com/ruoyi/business/mapper/HallInfoMapper.java",
    "platform-business/src/main/java/com/ruoyi/business/mapper/TerminalDeviceMapper.java",
    "platform-business/src/main/java/com/ruoyi/business/mapper/MediaResourceMapper.java",
    "platform-business/src/main/java/com/ruoyi/business/mapper/QueueInfoMapper.java",
    "platform-netty/pom.xml",
    "platform-netty/src/main/java/com/ruoyi/netty/server/NettyServer.java",
    "platform-netty/src/main/java/com/ruoyi/netty/handler/MessageHandler.java",
    "platform-netty/src/main/java/com/ruoyi/netty/protocol/ProtocolDecoder.java",
    "platform-netty/src/main/java/com/ruoyi/netty/protocol/ProtocolEncoder.java",
    "platform-netty/src/main/java/com/ruoyi/netty/service/DeviceStateService.java",
    "platform-netty/src/main/java/com/ruoyi/netty/websocket/WebSocketServer.java",
    "platform-system/pom.xml",
    "platform-system/src/main/java/com/ruoyi/system/controller/SysUserController.java",
    "platform-system/src/main/java/com/ruoyi/system/controller/SysRoleController.java",
    "platform-system/src/main/java/com/ruoyi/system/controller/SysMenuController.java",
    "platform-system/src/main/java/com/ruoyi/system/service/ISysUserService.java",
    "platform-system/src/main/java/com/ruoyi/system/service/impl/SysUserServiceImpl.java",
    "platform-system/src/main/java/com/ruoyi/system/domain/SysUser.java",
    "platform-system/src/main/java/com/ruoyi/system/domain/SysRole.java",
    "platform-system/src/main/java/com/ruoyi/system/domain/SysMenu.java",
    "platform-common/pom.xml",
    "platform-common/src/main/java/com/ruoyi/common/core/domain/AjaxResult.java",
    "platform-common/src/main/java/com/ruoyi/common/core/controller/BaseController.java",
    "platform-framework/pom.xml",
    "platform-framework/src/main/java/com/ruoyi/framework/config/SecurityConfig.java",
    "platform-framework/src/main/java/com/ruoyi/framework/web/service/TokenService.java",
    "platform-framework/src/main/java/com/ruoyi/framework/config/RedisConfig.java",
]

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


# ═══════════════════════════════════════════════════════════════════════
# 图表生成
# ═══════════════════════════════════════════════════════════════════════

def run_mermaid(name, mmd_content):
    os.makedirs(DIAGRAMS_DIR, exist_ok=True)
    mmd_path = os.path.join(DIAGRAMS_DIR, f"{name}.mmd")
    png_path = os.path.join(DIAGRAMS_DIR, f"{name}.png")
    with open(mmd_path, "w", encoding="utf-8") as f:
        f.write(mmd_content)
    result = subprocess.run(
        [MMDC, "-i", mmd_path, "-o", png_path, "-b", "white", "-w", "1200"],
        capture_output=True, text=True, timeout=60
    )
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
    result = subprocess.run(
        [DOT, "-Tpng", "-Gdpi=150", dot_path, f"-o{png_path}"],
        capture_output=True, text=True, timeout=30
    )
    if result.returncode == 0 and os.path.exists(png_path):
        print(f"  [OK] {name}.png")
    else:
        print(f"  [FAIL] {name}: {result.stderr[:200]}")
    return png_path


def generate_diagrams():
    """生成全部10张图表"""
    print("\n=== 生成图表 ===")

    # 01 系统架构图 (Graphviz)
    # 注意：subgraph label 必须用英文，否则中文乱码
    dot_arch = f'''digraph architecture {{
    rankdir=TB;
    node [shape=box, style="rounded,filled", fillcolor="#E8F4FD", fontname="Microsoft YaHei", fontsize=11];
    edge [fontname="Microsoft YaHei", fontsize=9];

    subgraph cluster_client {{ label="Client Layer"; style="rounded,dashed"; fillcolor="#F5F5F5";
        browser [label="浏览器 (Chrome/Edge)"];
    }}
    subgraph cluster_proxy {{ label="Reverse Proxy"; style="rounded,dashed"; fillcolor="#FFF3E0";
        nginx [label="Nginx"];
    }}
    subgraph cluster_app {{ label="Application Layer"; style="rounded,dashed"; fillcolor="#E8F5E9";
        backend [label="Spring Boot (platform-admin)"];
        ws [label="WebSocket Server"];
        netty [label="Netty TCP Server"];
    }}
    subgraph cluster_biz {{ label="Business Modules"; style="rounded,dashed"; fillcolor="#E3F2FD";
        hall [label="营业厅管理"];
        terminal [label="终端设备管理"];
        media [label="多媒体资源管理"];
        queue [label="排队叫号管理"];
    }}
    subgraph cluster_data {{ label="Data Layer"; style="rounded,dashed"; fillcolor="#FCE4EC";
        mysql [label="MySQL 5.7"];
        redis [label="Redis 5.0"];
    }}
    subgraph cluster_edge {{ label="Edge Devices"; style="rounded,dashed"; fillcolor="#F3E5F5";
        gateway [label="边缘代理网关"];
        device [label="终端设备"];
    }}

    browser -> nginx;
    nginx -> backend;
    nginx -> ws;
    backend -> hall;
    backend -> terminal;
    backend -> media;
    backend -> queue;
    backend -> mysql;
    backend -> redis;
    netty -> gateway;
    gateway -> device;
}}'''
    run_graphviz("01_system_architecture", dot_arch)

    # 02 模块分解图 (Graphviz)
    dot_mod = '''digraph modules {
    rankdir=TB;
    node [shape=box, style="rounded,filled", fontname="Microsoft YaHei", fontsize=10];

    subgraph cluster_front { label="Frontend (Vue.js)"; style="rounded"; fillcolor="#E3F2FD"; color="#1565C0";
        login [label="Login Module" fillcolor="#BBDEFB"];
        hall_fe [label="Hall Management" fillcolor="#BBDEFB"];
        terminal_fe [label="Terminal Management" fillcolor="#BBDEFB"];
        media_fe [label="Media Resource" fillcolor="#BBDEFB"];
        dashboard [label="Dashboard" fillcolor="#90CAF9"];
    }
    subgraph cluster_backend { label="Backend (Spring Boot)"; style="rounded"; fillcolor="#E8F5E9"; color="#2E7D32";
        ctrl [label="Controller Layer" fillcolor="#C8E6C9"];
        svc [label="Service Layer" fillcolor="#A5D6A7"];
        mapper [label="Mapper/DAO" fillcolor="#81C784"];
    }
    subgraph cluster_netty { label="Communication (Netty)"; style="rounded"; fillcolor="#FFF3E0"; color="#E65100";
        tcp [label="TCP Server" fillcolor="#FFE0B2"];
        proto [label="Protocol Codec" fillcolor="#FFCC80"];
        ws_mod [label="WebSocket" fillcolor="#FFB74D"];
    }

    login -> ctrl;
    hall_fe -> ctrl;
    ctrl -> svc;
    svc -> mapper;
    tcp -> proto;
    ws_mod -> svc;
}'''
    run_graphviz("02_module_decomposition", dot_mod)

    # 03 登录流程 (Mermaid)
    run_mermaid("03_login_flow", """flowchart TD
    A[用户打开浏览器] --> B[显示登录页面]
    B --> C[输入账号密码]
    C --> D{验证码校验}
    D -->|失败| E[提示验证码错误]
    E --> C
    D -->|成功| F{账号密码校验}
    F -->|失败| G{失败次数>=5?}
    G -->|是| H[账号锁定30分钟]
    G -->|否| I[提示密码错误]
    I --> C
    F -->|成功| J[生成JWT Token]
    J --> K[记录登录日志]
    K --> L[跳转系统首页]
    L --> M[显示数据仪表盘]""")

    # 04 营业厅管理流程 (Mermaid)
    run_mermaid("04_hall_management", """flowchart TD
    A[进入营业厅管理] --> B[查看营业厅列表]
    B --> C{选择操作}
    C -->|新增| D[填写营业厅信息]
    D --> E[校验编号唯一性]
    E -->|重复| D
    E -->|通过| F[保存到数据库]
    C -->|编辑| G[加载营业厅详情]
    G --> H[修改字段]
    H --> I[更新数据库]
    C -->|删除| J[确认删除]
    J --> K[检查关联设备]
    K -->|有关联| L[提示不可删除]
    K -->|无关联| M[删除记录]
    C -->|查询| N[输入筛选条件]
    N --> O[返回匹配结果]""")

    # 05 终端设备管理流程 (Mermaid)
    run_mermaid("05_terminal_management", """flowchart TD
    A[进入设备管理] --> B[加载设备列表]
    B --> C[显示在线/离线状态]
    C --> D{选择操作}
    D -->|新增设备| E[填写设备信息]
    E --> F[绑定边缘网关]
    F --> G[保存设备记录]
    D -->|设备详情| H[查看运行参数]
    H --> I[查看安装位置]
    I --> J[查看配件状态]
    D -->|远程控制| K[发送控制指令]
    K --> L[通过Netty下发]
    L --> M[等待设备响应]
    M -->|成功| N[更新设备状态]
    M -->|超时| O[标记通信异常]
    D -->|退役| P[确认退役]
    P --> Q[更新生命周期状态]""")

    # 06 资源下发流程 (Mermaid)
    run_mermaid("06_resource_publish", """flowchart TD
    A[进入资源管理] --> B[上传素材]
    B --> C{文件类型}
    C -->|图片| D[压缩处理]
    C -->|视频| E[转码处理]
    D --> F[存储到服务器]
    E --> F
    F --> G[资源列表展示]
    G --> H[选择目标终端]
    H --> I[选择播放时段]
    I --> J[确认下发]
    J --> K[WebSocket推送指令]
    K --> L[终端接收并播放]
    L --> M[上报播放状态]
    M --> N[更新资源使用统计]""")

    # 07 排队叫号流程 (Mermaid)
    run_mermaid("07_queue_management", """flowchart TD
    A[排队叫号终端] --> B[用户取号]
    B --> C[取号数+1]
    C --> D[显示排队信息]
    D --> E{叫号操作}
    E -->|叫号| F[叫号数+1]
    F --> G[显示下一号码]
    G --> H[语音播报]
    E -->|重叫| I[重新播报当前号码]
    E -->|过号| J[标记过号]
    J --> K[排队数-1]
    L[数据上报] --> M[实时同步到平台]
    M --> N[更新监控大屏]
    O[每日零点] --> P[自动重置计数器]""")

    # 08 Netty通信时序图 (Mermaid)
    run_mermaid("08_netty_sequence", """sequenceDiagram
    participant GW as 边缘网关
    participant NS as Netty Server
    participant PD as Protocol Decoder
    participant DS as DeviceStateService
    participant DB as MySQL
    participant WS as WebSocket
    participant FE as 前端页面

    GW->>NS: TCP连接建立
    NS->>PD: 解码二进制协议
    PD->>DS: 解析交易码+数据
    DS->>DB: 更新设备状态
    DS->>WS: 推送状态变更
    WS->>FE: 实时刷新设备列表

    GW->>NS: 心跳包(30s)
    NS->>DS: 更新在线时间
    DS->>WS: 心跳确认

    FE->>NS: 下发控制指令
    NS->>GW: TCP发送指令
    GW->>NS: 执行结果回传
    NS->>FE: WebSocket返回结果""")

    # 09 数据库ER图 (Mermaid)
    run_mermaid("09_database_er", """erDiagram
    HALL_INFO {
        bigint id PK
        varchar hall_code UK
        varchar hall_name
        varchar business_scope
        varchar level
        varchar business_hours
        varchar contact_person
        varchar contact_phone
        text address
        decimal longitude
        decimal latitude
    }
    TERMINAL_DEVICE {
        bigint id PK
        varchar device_code UK
        varchar device_type
        varchar model
        varchar install_location
        varchar mac_address
        varchar ip_address
        int online_status
        bigint hall_id FK
        bigint gateway_id FK
    }
    EDGE_GATEWAY {
        bigint id PK
        varchar gateway_code UK
        varchar gateway_name
        varchar ip_address
        int port
        int online_status
    }
    MEDIA_RESOURCE {
        bigint id PK
        varchar resource_name
        varchar resource_type
        varchar file_path
        bigint file_size
        varchar category
    }
    QUEUE_INFO {
        bigint id PK
        bigint terminal_id FK
        int take_count
        int call_count
        int queue_count
        int wait_count
        varchar queue_date
    }
    DEVICE_ACCESSORY {
        bigint id PK
        bigint device_id FK
        varchar accessory_type
        varchar status
        decimal temperature
        varchar mode
    }
    HALL_INFO ||--o{ TERMINAL_DEVICE : "has"
    EDGE_GATEWAY ||--o{ TERMINAL_DEVICE : "connects"
    TERMINAL_DEVICE ||--o{ QUEUE_INFO : "has"
    TERMINAL_DEVICE ||--o{ DEVICE_ACCESSORY : "has"
    MEDIA_RESOURCE }o--o{ TERMINAL_DEVICE : "deployed to" """)

    # 10 数据流图 (Mermaid)
    run_mermaid("10_data_flow", """flowchart LR
    subgraph Input
        A1[浏览器请求] --> B1[API接口]
        A2[边缘网关上报] --> B2[TCP协议]
        A3[终端设备状态] --> B3[WebSocket]
    end
    subgraph Process
        B1 --> C1[Controller]
        C1 --> C2[Service业务逻辑]
        C2 --> C3[Mapper数据访问]
        B2 --> C4[Netty协议解析]
        C4 --> C2
        B3 --> C5[WebSocket推送]
    end
    subgraph Storage
        C3 --> D1[(MySQL)]
        C2 --> D2[(Redis缓存)]
    end
    subgraph Output
        D1 --> E1[前端页面渲染]
        D2 --> E1
        C5 --> E2[实时状态大屏]
        C2 --> E3[报表统计]
    end""")

    print("图表生成完毕")


# ═══════════════════════════════════════════════════════════════════════
# 文档1：申请表 (.doc)
# ═══════════════════════════════════════════════════════════════════════

def generate_application_form():
    """
    复制模板 .doc 并替换字段值，保留原始格式。
    如果模板不存在，则从零创建（格式可能有偏差）。
    """
    print("\n=== 生成申请表 ===")
    output_path = os.path.join(OUTPUT_DIR, f"{SOFTWARE_NAME}.doc")

    if os.path.exists(TEMPLATE_DOC):
        # 方案A：复制模板再替换（推荐，格式一致）
        import shutil
        shutil.copy2(TEMPLATE_DOC, output_path)
        doc = Document(output_path)
        table = doc.tables[0]
        for idx, val in FILL_VALUES.items():
            if val == "":
                continue
            if idx < len(table.rows):
                cell = table.rows[idx].cells[1]
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.text = ""
                if cell.paragraphs:
                    cell.paragraphs[0].clear()
                    run = cell.paragraphs[0].add_run(val)
                    # 保持模板原始字体
                    set_run_font(run, "微软雅黑", 14)
        doc.save(output_path)
    else:
        # 方案B：从零创建（无模板时的后备方案）
        print(f"  [WARN] 模板不存在: {TEMPLATE_DOC}，从零创建")
        doc = Document()
        section = doc.sections[0]
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

        title = doc.add_paragraph()
        title.alignment = 1
        title.space_after = Pt(12)
        run = title.add_run("公司软件著作权信息登记采集表")
        set_run_font(run, "宋体", 16, bold=True)

        FIELD_LABELS = [
            "著作权人姓名或名称（必填）",
            "公司成立日期（必填）",
            "营业执照注册号、身份证号或事业单位法人证书编号（必填）",
            "公司地址（必填）",
            "软件名称（必填）",
            "版本号（必填）",
            "是否是升级版本（必填）",
            "软件简称（选填）",
            "软件分类（必填）",
            "开发完成日期（必填）",
            "发表状态（必填）",
            "首次发表时间（必填，若尚未发表，不填）",
            "首次发表地点（必填，若尚未发表，不填）",
            "开发的硬件环境（必填）",
            "运行的硬件环境（必填）",
            "开发该软件的操作系统（必填）",
            "软件开发环境/开发工具（必填）",
            "该软件的运行平台/操作系统（必填）",
            "软件运行支撑环境/支持软件（必填）",
            "编程语言（必填）",
            "源程序量（必填）",
            "开发目的（50字以内）（必填）",
            "面向领域/行业（必填）",
            "主要功能（必填）",
            "技术特点（必填）",
            "软件技术特点（必填）",
        ]
        table = doc.add_table(rows=26, cols=2)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i in range(26):
            cell_label = table.cell(i, 0)
            cell_value = table.cell(i, 1)
            cell_label.width = Cm(7)
            cell_value.width = Cm(10)
            cell_label.paragraphs[0].clear()
            r0 = cell_label.paragraphs[0].add_run(FIELD_LABELS[i])
            set_run_font(r0)
            cell_value.paragraphs[0].clear()
            val = FILL_VALUES.get(i, "")
            r1 = cell_value.paragraphs[0].add_run(val)
            set_run_font(r1)
            set_cell_shading(cell_label, "D9E2F3")
        doc.save(output_path)

    print(f"  [OK] {output_path}")


# ═══════════════════════════════════════════════════════════════════════
# 文档2：合并源代码 (.docx)
# ═══════════════════════════════════════════════════════════════════════

def generate_merged_code():
    """合并源代码为60页文档（前30+后30），每50行一页"""
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

    # 每50行一页，取前30+后30页
    PAGE_SIZE = 50
    pages = [all_lines[i:i + PAGE_SIZE] for i in range(0, len(all_lines), PAGE_SIZE)]
    if len(pages) <= 60:
        selected_pages = pages
    else:
        selected_pages = pages[:30] + pages[-30:]

    # 生成 docx
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # 页眉
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
# 文档3：软件设计说明书 (.docx)
# ═══════════════════════════════════════════════════════════════════════

def generate_design_spec():
    """生成软件设计说明书，嵌入图表"""
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
    run = title.add_run(f"{SOFTWARE_NAME}")
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

    # 目录页
    add_heading_styled(doc, "目录", 1)
    toc_items = [
        "一、引言",
        "二、总体设计",
        "三、接口设计",
        "四、模块设计",
        "五、运行设计",
        "六、出错处理设计",
        "七、数据库设计",
    ]
    for item in toc_items:
        add_para(doc, item, "宋体", 12, indent=0)
    doc.add_page_break()

    # ─── 一、引言 ───────────────────────────────────────────────
    add_heading_styled(doc, "一、引言", 1)

    add_heading_styled(doc, "1.1 编写目的", 2)
    add_para(doc, f"本文档是《{SOFTWARE_NAME}》的软件设计说明书，旨在详细描述系统的总体设计、模块划分、接口设计、算法逻辑和运行方案，为软件著作权登记提供技术文档支撑。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "1.2 项目背景", 2)
    add_para(doc, "随着电力行业营业厅数字化转型的推进，营业厅内电子设备数量不断增加，设备管理、运维和媒体资源发布面临巨大挑战。本项目旨在构建一套基于边缘代理网关的营业厅运营管理平台，实现终端设备的远程接入、实时状态监控、媒体资源集中管控和运维管理的数字化。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "1.3 术语定义", 2)
    terms = [
        ("边缘代理网关", "部署在营业厅现场的轻量级网关设备，负责终端设备的数据采集、协议转换和状态上报"),
        ("Netty", "Java NIO客户端/服务端框架，用于构建高性能TCP长连接通信服务"),
        ("JWT", "JSON Web Token，用于分布式环境下的用户身份认证"),
        ("RBAC", "基于角色的访问控制模型，用于系统权限管理"),
    ]
    for term, desc in terms:
        add_para(doc, f"{term}：{desc}", "宋体", 12, indent=0.74)

    doc.add_page_break()

    # ─── 二、总体设计 ─────────────────────────────────────────────
    add_heading_styled(doc, "二、总体设计", 1)

    add_heading_styled(doc, "2.1 系统架构", 2)
    add_para(doc, f"本系统采用前后端分离的B/S架构，基于Spring Boot + Vue.js技术栈构建。系统分为客户端层、反向代理层、应用服务层、业务模块层、数据层和边缘设备层六个层次。", "宋体", 12, indent=0.74)
    add_para(doc, "系统架构图如下：", "宋体", 12, indent=0.74)
    _add_image(doc, "01_system_architecture")

    add_heading_styled(doc, "2.2 技术栈", 2)
    tech_rows = [
        ("后端框架", "Spring Boot 2.5.15 + Spring Security + MyBatis"),
        ("前端框架", "Vue.js 2.6.12 + Element UI 2.15.13"),
        ("数据库", "MySQL 5.7"),
        ("缓存", "Redis 5.0"),
        ("通信", "Netty 4.x TCP长连接 + WebSocket"),
        ("认证", "JWT Token"),
        ("构建", "Maven 多模块 + Webpack"),
    ]
    add_table_from_rows(doc, ["层次", "技术选型"], tech_rows, [4, 10])

    add_heading_styled(doc, "2.3 模块划分", 2)
    add_para(doc, "系统按功能划分为前端模块、后端业务模块和通信模块三大板块。模块分解图如下：", "宋体", 12, indent=0.74)
    _add_image(doc, "02_module_decomposition")

    doc.add_page_break()

    # ─── 三、接口设计 ─────────────────────────────────────────────
    add_heading_styled(doc, "三、接口设计", 1)

    add_heading_styled(doc, "3.1 REST API接口", 2)
    add_para(doc, "系统后端提供RESTful风格的HTTP接口，前端通过Axios调用。主要接口模块如下：", "宋体", 12, indent=0.74)
    api_rows = [
        ("/api/hall/**", "营业厅管理", "CRUD操作、分页查询"),
        ("/api/terminal/**", "终端设备管理", "设备增删改查、状态监控"),
        ("/api/media/**", "多媒体资源", "资源上传、下载、下发"),
        ("/api/queue/**", "排队叫号", "取号、叫号、数据查询"),
        ("/api/accessory/**", "设备配件", "配件状态查询、参数设置"),
        ("/api/system/**", "系统管理", "用户、角色、菜单管理"),
    ]
    add_table_from_rows(doc, ["接口路径", "模块", "功能"], api_rows, [4, 3, 6])

    add_heading_styled(doc, "3.2 WebSocket接口", 2)
    add_para(doc, "系统使用WebSocket实现服务端向客户端的实时推送，主要场景包括设备状态变更通知、排队叫号数据更新等。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "3.3 Netty TCP接口", 2)
    add_para(doc, "边缘代理网关通过TCP长连接与平台通信，采用自定义二进制协议，包含消息头（长度+交易码）和消息体。Netty Server监听端口30002。", "宋体", 12, indent=0.74)

    doc.add_page_break()

    # ─── 四、模块设计 ─────────────────────────────────────────────
    add_heading_styled(doc, "四、模块设计", 1)

    # 4.1 用户登录模块
    add_heading_styled(doc, "4.1 用户登录模块", 2)
    add_heading_styled(doc, "4.1.1 功能说明", 3)
    add_para(doc, "用户登录模块负责系统身份认证，支持账号密码登录、验证码校验、登录失败次数锁定、JWT Token生成等功能。基于Spring Security框架实现，密码采用BCrypt加密存储。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "4.1.2 处理流程图", 3)
    _add_image(doc, "03_login_flow")
    add_heading_styled(doc, "4.1.3 关键函数", 3)
    login_funcs = [
        ("login()", "SysLoginController", "处理登录请求，校验验证码和账号密码"),
        ("authenticate()", "AuthenticationManager", "Spring Security认证管理器"),
        ("createToken()", "TokenService", "生成JWT Token"),
        ("recordLoginLog()", "SysLoginService", "记录登录日志"),
    ]
    add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], login_funcs, [3, 4, 7])
    add_heading_styled(doc, "4.1.4 算法说明", 3)
    add_para(doc, "登录失败次数锁定算法：维护计数器记录连续失败次数，每失败一次计数器加1并设置30分钟过期时间。当计数器值达到5时，返回账号锁定提示。登录成功时清零计数器。", "宋体", 12, indent=0.74)

    # 4.2 营业厅管理模块
    add_heading_styled(doc, "4.2 营业厅管理模块", 2)
    add_heading_styled(doc, "4.2.1 功能说明", 3)
    add_para(doc, "营业厅管理模块维护营业厅基本信息，包括编号、名称、营业范围、级别、营业时间、负责人、联系电话、详细地址和经纬度坐标。支持新增、修改、删除和查询操作，编号具有唯一性约束。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "4.2.2 处理流程图", 3)
    _add_image(doc, "04_hall_management")
    add_heading_styled(doc, "4.2.3 关键函数", 3)
    hall_funcs = [
        ("list()", "HallInfoController", "分页查询营业厅列表"),
        ("add()", "HallInfoController", "新增营业厅信息"),
        ("edit()", "HallInfoController", "修改营业厅信息"),
        ("remove()", "HallInfoController", "删除营业厅（校验关联设备）"),
        ("checkHallCodeUnique()", "HallInfoServiceImpl", "校验营业厅编号唯一性"),
    ]
    add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], hall_funcs, [4, 4, 6])
    add_heading_styled(doc, "4.2.4 算法说明", 3)
    add_para(doc, "唯一性校验算法：新增或修改营业厅时，根据编号查询数据库是否存在相同编号的记录。新增时若已存在则返回错误；修改时排除自身ID后查询，若存在则返回错误。", "宋体", 12, indent=0.74)

    # 4.3 终端设备管理模块
    add_heading_styled(doc, "4.3 终端设备管理模块", 2)
    add_heading_styled(doc, "4.3.1 功能说明", 3)
    add_para(doc, "终端设备管理模块管理营业厅内所有电子终端设备的全生命周期，包括设备编号、类型、型号、安装位置、购置日期、MAC地址、IP地址等信息。支持设备在线/离线状态实时监控，记录设备从采购入库、安装运行到报废退役的完整生命周期。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "4.3.2 处理流程图", 3)
    _add_image(doc, "05_terminal_management")
    add_heading_styled(doc, "4.3.3 关键函数", 3)
    term_funcs = [
        ("list()", "TerminalDeviceController", "分页查询设备列表"),
        ("add()", "TerminalDeviceController", "新增设备信息"),
        ("getStatus()", "DeviceStateService", "获取设备实时在线状态"),
        ("sendCommand()", "NettyServer", "通过Netty下发控制指令"),
        ("updateLifecycle()", "TerminalDeviceServiceImpl", "更新设备生命周期状态"),
    ]
    add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], term_funcs, [4, 4, 6])
    add_heading_styled(doc, "4.3.4 算法说明", 3)
    add_para(doc, "设备状态检测算法：Netty Server维护设备连接映射表（deviceId -> Channel），每30秒检测一次Channel活跃状态。若Channel不活跃则标记设备离线并通知WebSocket推送前端更新。", "宋体", 12, indent=0.74)

    # 4.4 多媒体资源管理模块
    add_heading_styled(doc, "4.4 多媒体资源管理模块", 2)
    add_heading_styled(doc, "4.4.1 功能说明", 3)
    add_para(doc, "多媒体资源管理模块管理营业厅广告机等终端的多媒体资源，支持图片和视频素材的上传、分类、预览，可将资源远程下发到指定终端设备进行播放。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "4.4.2 处理流程图", 3)
    _add_image(doc, "06_resource_publish")
    add_heading_styled(doc, "4.4.3 关键函数", 3)
    media_funcs = [
        ("upload()", "MediaResourceController", "上传多媒体资源文件"),
        ("list()", "MediaResourceController", "分页查询资源列表"),
        ("publish()", "MediaResourceServiceImpl", "下发资源到指定终端"),
        ("getPlayStatus()", "MediaResourceServiceImpl", "查询终端播放状态"),
    ]
    add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], media_funcs, [4, 4, 6])
    add_heading_styled(doc, "4.4.4 算法说明", 3)
    add_para(doc, "资源下发算法：选择目标终端后，平台通过WebSocket向终端推送资源下发指令，包含资源URL、播放时段和播放模式。终端接收指令后下载资源并按配置播放，定时上报播放状态。", "宋体", 12, indent=0.74)

    # 4.5 排队叫号管理模块
    add_heading_styled(doc, "4.5 排队叫号管理模块", 2)
    add_heading_styled(doc, "4.5.1 功能说明", 3)
    add_para(doc, "排队叫号管理模块实时获取终端排队叫号数据，包括取号数、叫号数、排队数和等待数，支持按日期自动重置计数。通过WebSocket实现数据实时同步到管理平台。", "宋体", 12, indent=0.74)
    add_heading_styled(doc, "4.5.2 处理流程图", 3)
    _add_image(doc, "07_queue_management")
    add_heading_styled(doc, "4.5.3 关键函数", 3)
    queue_funcs = [
        ("takeNumber()", "QueueInfoServiceImpl", "用户取号，取号数+1"),
        ("callNumber()", "QueueInfoServiceImpl", "叫号操作，叫号数+1"),
        ("getData()", "QueueInfoController", "查询实时排队数据"),
        ("resetDaily()", "QueueInfoServiceImpl", "每日零点自动重置计数器"),
    ]
    add_table_from_rows(doc, ["函数名", "所属类", "功能说明"], queue_funcs, [4, 4, 6])
    add_heading_styled(doc, "4.5.4 算法说明", 3)
    add_para(doc, "每日重置算法：使用Redis的EXPIRE机制设置计数器过期时间。取号时检查当日计数器是否存在，不存在则创建并设置为当天剩余秒数的过期时间。每日零点计数器自动过期，新的一天从0开始计数。", "宋体", 12, indent=0.74)

    doc.add_page_break()

    # ─── 五、运行设计 ─────────────────────────────────────────────
    add_heading_styled(doc, "五、运行设计", 1)

    add_heading_styled(doc, "5.1 运行流程", 2)
    add_para(doc, "用户通过浏览器访问平台，经过Nginx反向代理到达Spring Boot应用。登录成功后进入系统首页，通过左侧导航菜单访问各业务模块。系统通过Netty TCP Server与边缘网关保持长连接，实时采集设备状态。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "5.2 通信时序", 2)
    add_para(doc, "边缘网关与平台的核心通信时序如下：", "宋体", 12, indent=0.74)
    _add_image(doc, "08_netty_sequence")

    add_heading_styled(doc, "5.3 数据流", 2)
    add_para(doc, "系统的数据流路径：输入层（浏览器请求、网关上报、设备状态）→ 处理层（Controller → Service → Mapper / Netty协议解析）→ 存储层（MySQL + Redis）→ 输出层（页面渲染、实时大屏、报表统计）。", "宋体", 12, indent=0.74)
    _add_image(doc, "10_data_flow")

    doc.add_page_break()

    # ─── 六、出错处理设计 ───────────────────────────────────────────
    add_heading_styled(doc, "六、出错处理设计", 1)

    add_heading_styled(doc, "6.1 接口异常处理", 2)
    add_para(doc, "REST API统一使用AjaxResult封装返回结果，包含code、msg、data三个字段。Controller层通过try-catch捕获业务异常，Service层抛出自定义BusinessException。全局异常处理器@ExceptionHandler捕获未处理异常，返回统一错误格式。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "6.2 通信异常处理", 2)
    add_para(doc, "Netty TCP通信异常通过ChannelInactive事件检测设备断连，自动更新设备状态为离线。WebSocket断连后前端自动重连（指数退避策略）。消息发送失败时记录日志并触发告警。", "宋体", 12, indent=0.74)

    add_heading_styled(doc, "6.3 数据库异常处理", 2)
    add_para(doc, "数据库操作异常由MyBatis统一捕获，记录错误日志。涉及事务的操作使用@Transactional注解，异常时自动回滚。连接池（Druid）监控慢SQL和连接泄漏。", "宋体", 12, indent=0.74)

    doc.add_page_break()

    # ─── 七、数据库设计 ─────────────────────────────────────────────
    add_heading_styled(doc, "七、数据库设计", 1)

    add_heading_styled(doc, "7.1 ER关系图", 2)
    _add_image(doc, "09_database_er")

    add_heading_styled(doc, "7.2 核心表结构", 2)
    add_para(doc, "系统核心数据表包括：", "宋体", 12, indent=0.74)
    db_rows = [
        ("hall_info", "营业厅信息表", "id, hall_code, hall_name, address, contact_phone, longitude, latitude"),
        ("terminal_device", "终端设备表", "id, device_code, device_type, hall_id, gateway_id, online_status"),
        ("edge_gateway", "边缘网关表", "id, gateway_code, ip_address, port, online_status"),
        ("media_resource", "多媒体资源表", "id, resource_name, resource_type, file_path, file_size"),
        ("queue_info", "排队叫号表", "id, terminal_id, take_count, call_count, queue_date"),
        ("device_accessory", "设备配件表", "id, device_id, accessory_type, status, temperature"),
        ("sys_user", "系统用户表", "user_id, user_name, password, status, dept_id"),
        ("sys_role", "角色表", "role_id, role_name, role_key, status"),
        ("sys_menu", "菜单表", "menu_id, menu_name, parent_id, path, component"),
    ]
    add_table_from_rows(doc, ["表名", "说明", "核心字段"], db_rows, [3, 3, 8])

    # 保存
    doc.save(output_path)
    print(f"  [OK] {output_path}")


def _add_image(doc, diagram_name):
    """嵌入图表，找不到时输出提示"""
    img_path = os.path.join(DIAGRAMS_DIR, f"{diagram_name}.png")
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=Inches(5.5))
    else:
        add_para(doc, f"[图表缺失: {diagram_name}.png]", "宋体", 10)


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
    print(f"  4. diagrams/ (10张PNG)")


if __name__ == "__main__":
    main()
