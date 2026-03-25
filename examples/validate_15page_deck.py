# -*- coding: utf-8 -*-
"""
验证脚本：pptx-creator-agent 改进成果汇报（16 页）

用途：验证分块生成工作流
  - 先定义完整 slide plan（16 页）
  - builders 拆分为 2 批（每批 8 页）
  - 每批执行后打印进度
  - 最后调用 sanitize_script 演示代码卫生校验

运行:
    python validate_15page_deck.py
"""

import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))

from pptx_helpers import *


# ═══════════════════════════════════════════════════
# Slide Plan（规划表）
# ═══════════════════════════════════════════════════

SLIDE_PLAN = [
    (1,  '封面',              '全屏标题',     'pptx-creator-agent 改进成果汇报'),
    (2,  '目录',              '目录',         '5 章节导览'),
    (3,  '改进前现状',        '表格',         '4 次实战成功率数据'),
    (4,  '核心问题分析',      '混合布局',     '5 大失败原因'),
    (5,  '致命缺陷详解',      '代码块+说明',  '响应长度限制根因'),
    (6,  '改进方案总览',      '信息框组合',   '3 条核心改进路径'),
    (7,  'pptx_helpers 设计', '表格',         '12 个元素级 Helper API'),
    (8,  '分块生成策略',      '步骤展示',     '5 步标准生成流程'),
    (9,  '代码卫生检查',      '代码块',       'sanitize_script 实现'),
    (10, '核心 API 速查',     '表格',         '常用参数参考'),
    (11, '配色方案对比',      '表格',         '3 套内置配色方案'),
    (12, 'Builder 模式示例',  '代码大图',     '推荐的脚本骨架'),
    (13, '改进前后对比',      '对比表格',     '行数 / 重复代码 / 成功率'),
    (14, '成果数据',          '数据展示',     '脚本体积减少 40%'),
    (15, '关键结论',          '混合布局',     '5 条实战结论'),
    (16, '后续计划 / Q&A',    '全屏收束',     '下一阶段目标'),
]

TOTAL = len(SLIDE_PLAN)


def print_slide_plan():
    print('Slide plan:')
    for page, title, layout, content in SLIDE_PLAN:
        print(f'  {page:02d}. {title:<20} | {layout:<12} | {content}')
    print()


def run_builder_batch(prs, C, batch_name, builders):
    print(f'Running {batch_name} ({len(builders)} slides)...')
    for builder in builders:
        builder(prs, C)
    print(f'  done — {len(prs.slides)} slides total so far')


# ═══════════════════════════════════════════════════
# Batch 1: Slides 01–08  背景与方案
# ═══════════════════════════════════════════════════

def build_slide_01_cover(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, 0, 0, SLIDE_W, SLIDE_H, C['navy_dark'])
    add_bg(s, 0, Emu(2200000), SLIDE_W, Emu(2700000), C['navy'])
    add_bg(s, 0, Emu(2200000), SLIDE_W, Emu(55000), C['teal'])
    add_bg(s, 0, Emu(4845000), SLIDE_W, Emu(55000), C['teal'])
    tb(s, Emu(800000), Emu(2450000), Emu(10500000), Emu(700000),
       'pptx-creator-agent 改进成果汇报',
       sz=32, b=True, c=C['white'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(3250000), Emu(10500000), Emu(450000),
       '规划先行  |  Helper 复用  |  分块生成  |  代码卫生校验',
       sz=13, c=C['gray_mid'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(4100000), Emu(10500000), Emu(350000),
       '2026-03-19   基于 4 次 PPTX 生成实战经验',
       sz=11, c=C['teal'], al=PP_ALIGN.CENTER, C=C)
    sn(s, 1, TOTAL, C=C)


def build_slide_02_agenda(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '目录', C=C)
    chapters = [
        ('一', '改进前现状',   '4 次实战成功率 + 主要失败模式',           C['red']),
        ('二', '改进方案',     '3 条核心改进路径',                         C['teal']),
        ('三', '技术实现',     'pptx_helpers + 分块策略 + 代码卫生',       C['navy']),
        ('四', '改进成果',     '前后对比数据',                             C['green']),
        ('五', '结论与展望',   '关键结论 + 下一步',                        C['amber']),
    ]
    for i, (num, title, desc, clr) in enumerate(chapters):
        y = Emu(1100000 + i * 1000000)
        badge(s, Emu(400000), y, Emu(520000), Emu(450000), num, clr, C=C)
        tb(s, Emu(1100000), y + Emu(30000), Emu(4000000), Emu(390000),
           title, sz=20, b=True, c=clr, C=C)
        tb(s, Emu(5300000), y + Emu(80000), Emu(6500000), Emu(310000),
           desc, sz=12, c=C['gray_text'], C=C)
    sn(s, 2, TOTAL, C=C)


def build_slide_03_status(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '改进前现状', 'pptx-creator 实战成功率', C=C)
    act_badge(s, Emu(350000), Emu(870000), '一', '改进前现状', C['red'], C=C)

    hdr = ['#', '文件', '页数', '生成方式', '结果']
    data = [
        ['1', '项目工具链综述',       '15', 'pptx-creator',                   '✅ 成功'],
        ['2', '测试场景分析报告',     '23', 'pptx-creator → 手写脚本',         '❌ 失败'],
        ['3', '市场调研报告',         '32', 'pptx-creator → 手写脚本',         '❌ 失败'],
        ['4', '问题排查汇报',         '28', '直接手写脚本',                     '❌ 跳过'],
    ]
    cw = [Emu(450000), Emu(2800000), Emu(700000), Emu(3800000), Emu(1200000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(8950000), Emu(3200000), 5, 5, cw)
    ft(t, hdr, data, fs=11, hfs=12, C=C)

    result_colors = [C['green'], C['red'], C['red'], C['red']]
    result_texts = ['✅ 成功', '❌ 失败', '❌ 失败', '❌ 跳过']
    for row in range(1, 5):
        sc(t.cell(row, 4), result_texts[row - 1], sz=11,
           fc=result_colors[row - 1], C=C)

    add_bg(s, Emu(9500000), Emu(1500000), Emu(2400000), Emu(3200000), C['light_red'])
    tb(s, Emu(9650000), Emu(1650000), Emu(2100000), Emu(450000),
       '成功率', sz=14, b=True, c=C['red'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(9650000), Emu(2200000), Emu(2100000), Emu(800000),
       '1/4', sz=52, b=True, c=C['red'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(9650000), Emu(3100000), Emu(2100000), Emu(380000),
       '25%', sz=22, b=True, c=C['red'], al=PP_ALIGN.CENTER, C=C)

    hbox(s, Emu(350000), Emu(4900000), Emu(11500000), Emu(1400000),
         '根本原因',
         ['每次尝试：读取源文档 → 理解内容 → 规划幻灯片 → 生成完整脚本（800~1300 行）',
          '当源文档超过 200~300 行时，prompt 超出上下文窗口，生成被截断或直接失败'],
         bg=C['light_red'], tc=C['red'], C=C)
    sn(s, 3, TOTAL, C=C)


def build_slide_04_problems(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '核心问题分析', '5 大失败模式', C=C)
    act_badge(s, Emu(350000), Emu(870000), '一', '改进前现状', C['red'], C=C)

    problems = [
        ('P1', '响应长度硬限制',       '脚本 800~1300 行，超出单次输出能力，是 3/4 次失败的直接原因',    C['red']),
        ('P2', '无法处理大量源材料',   '1872 行源文档无预处理机制，prompt 溢出',                          C['amber']),
        ('P3', '设计系统无复用',       '每次重定义 150 行 helper + 配色方案，4 个脚本 = 600 行重复代码',   C['orange']),
        ('P4', 'Unicode 引号 SyntaxError', '生成代码混入弯引号，每次需要手动修复才能运行',                C['teal']),
        ('P5', '无分块/增量生成能力',  '失败后只能全部重来，无断点续写机制',                             C['navy']),
    ]
    for i, (num, title, desc, clr) in enumerate(problems):
        y = Emu(1450000 + i * 950000)
        badge(s, Emu(350000), y, Emu(520000), Emu(390000), num, clr, C=C)
        tb(s, Emu(1050000), y + Emu(15000), Emu(3200000), Emu(360000),
           title, sz=14, b=True, c=clr, C=C)
        tb(s, Emu(4450000), y + Emu(35000), Emu(7400000), Emu(330000),
           desc, sz=11, c=C['black'], C=C)
    sn(s, 4, TOTAL, C=C)


def build_slide_05_fatal_issue(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '致命缺陷详解', '响应长度硬限制 — 根因与数据', C=C)
    act_badge(s, Emu(350000), Emu(870000), '一', '改进前现状', C['red'], C=C)

    tb(s, Emu(350000), Emu(1500000), Emu(5200000), Emu(350000),
       'Prompt 组成（典型 22 页 deck）', sz=13, b=True, c=C['navy'], C=C)
    ml(s, Emu(350000), Emu(1900000), Emu(5200000), Emu(2200000),
       ['agent 系统指令         ~500 tokens',
        '用户源文档（MD 文件）   ~3000–8000 tokens',
        '规划 + 代码输出         ~3000–6000 tokens',
        '─────────────────────────────',
        '合计超出上下文窗口上限'],
       sz=11, c=C['black'], fn=C['font_code'], C=C)

    tb(s, Emu(5800000), Emu(1500000), Emu(5900000), Emu(350000),
       '典型错误输出', sz=13, b=True, c=C['navy'], C=C)
    code(s, Emu(5750000), Emu(1900000), Emu(6050000), Emu(2200000),
         ['# 失败模式 A：直接报错',
          'Error: prompt too long',
          '',
          '# 失败模式 B：生成被截断',
          'def build_slide_22(prs, C):',
          '    s = prs.slides.add_slide(...)',
          '    # [output truncated at token limit]'],
         sz=11, C=C)

    hbox(s, Emu(350000), Emu(4400000), Emu(11500000), Emu(1900000),
         '解决思路',
         ['① 源文档 > 300 行时，先提取摘要（节标题 + 要点 + 关键数据块）',
          '② 生成脚本骨架（imports + create_prs + builders 列表 + main），暂不填充函数体',
          '③ 按批次生成 slide builder 函数（每批 ≤8 页），追加到同一文件',
          '④ 复用 pptx_helpers.py，消除每个脚本里 150 行重复 helper 定义'],
         bg=C['light_blue'], tc=C['navy'], C=C)
    sn(s, 5, TOTAL, C=C)


def build_slide_06_solutions(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '改进方案总览', '3 条核心改进路径', C=C)
    act_badge(s, Emu(350000), Emu(870000), '二', '改进方案', C['teal'], C=C)

    hbox(s, Emu(350000), Emu(1500000), Emu(3600000), Emu(4600000),
         '① Helper 复用',
         ['创建 pptx_helpers.py',
          '12 个元素级函数',
          '3 套内置配色方案',
          'from pptx_helpers import *',
          '脚本减少 40% 行数',
          '彻底消除重复代码'],
         bg=C['light_blue'], tc=C['navy'], C=C)
    hbox(s, Emu(4200000), Emu(1500000), Emu(3600000), Emu(4600000),
         '② 分块生成',
         ['大纲规划先行',
          '脚本骨架与 builders 分离',
          '每批 ≤8 页',
          '追加到同一文件',
          '批次完成后语法校验',
          '失败从断点恢复'],
         bg=C['light_green'], tc=C['green'], C=C)
    hbox(s, Emu(8050000), Emu(1500000), Emu(3700000), Emu(4600000),
         '③ 代码卫生',
         ['ast.parse 语法验证',
          'Unicode 弯引号检测',
          'sanitize_script() 自动修复',
          '执行前强制校验',
          '优先使用 ASCII 引号',
          '消除 SyntaxError'],
         bg=C['light_amber'], tc=C['amber'], C=C)
    sn(s, 6, TOTAL, C=C)


def build_slide_07_helpers_design(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, 'pptx_helpers.py 设计', '12 个元素级 Helper 函数', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    hdr = ['函数', '用途', '粒度', '使用频次']
    data = [
        ['add_bg()',            '纯色矩形背景块',             '元素级', '极高（每页 3-5 次）'],
        ['tb()',                '单行/多行文本框',             '元素级', '极高（每页 2-4 次）'],
        ['ml()',                '多行文本 + 项目符号',         '元素级', '高'],
        ['rl()',                '富文本列表（每行独立样式）',  '元素级', '中'],
        ['ct() + ft()',         '创建表格 + 填充数据',         '元素级', '高'],
        ['sc()',                '设置单元格内容和样式',        '单元格级', '高'],
        ['code()',              '深色背景代码块',              '元素级', '中（技术文档）'],
        ['hbox()',              '带标题的高亮信息框',          '元素级', '高'],
        ['badge() / act_badge()','序号徽章 / 章节标记条',     '元素级', '中'],
        ['bar()',               '顶部装饰条 + 标题',          '元素级', '极高（几乎每页）'],
        ['sn()',                '右下角页码',                  '元素级', '每页'],
        ['sanitize_script()',   'Unicode 修复 + 语法校验',    '工具函数', '执行前'],
    ]
    cw = [Emu(2200000), Emu(3800000), Emu(1500000), Emu(2500000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(10000000), Emu(5000000), 13, 4, cw)
    ft(t, hdr, data, fs=10, hfs=11, C=C)
    sn(s, 7, TOTAL, C=C)


def build_slide_08_chunking(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '分块生成策略', '5 步标准生成流程', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    steps = [
        ('摘要',  '源文档 > 300 行\n先提取摘要\n保留标题+要点+数据',   C['teal']),
        ('规划',  '输出 slide plan\n页码/标题/布局\n/关键内容',         C['navy']),
        ('骨架',  'imports\ncreate_prs()\nbuilders 列表\nmain()',        C['amber']),
        ('批次',  '每批生成\n≤8 个 builders\n追加到同一脚本',           C['green']),
        ('校验',  'ast.parse\nsanitize_script\n统一执行验证',            C['red']),
    ]
    for i, (title, desc, clr) in enumerate(steps):
        x = Emu(350000 + i * 2340000)
        add_bg(s, x, Emu(1500000), Emu(2150000), Emu(3800000), clr)
        tb(s, x + Emu(100000), Emu(1620000), Emu(1950000), Emu(350000),
           f'Step {i + 1}', sz=11, c=C['white'], fn=C['font_code'], C=C)
        tb(s, x + Emu(100000), Emu(2000000), Emu(1950000), Emu(420000),
           title, sz=18, b=True, c=C['white'], C=C)
        ml(s, x + Emu(80000), Emu(2550000), Emu(2020000), Emu(2500000),
           desc.split('\n'), sz=10, c=C['white'], ls=1.3, C=C)

    tb(s, Emu(350000), Emu(5500000), Emu(11500000), Emu(600000),
       '每批之间独立追加 ·  失败时从最后完成的 slide builder 断点恢复 ·  不需要整份重来',
       sz=11, c=C['gray_text'], al=PP_ALIGN.CENTER, C=C)
    sn(s, 8, TOTAL, C=C)


# ═══════════════════════════════════════════════════
# Batch 2: Slides 09–16  实现细节与成果
# ═══════════════════════════════════════════════════

def build_slide_09_hygiene(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '代码卫生检查', 'sanitize_script() 实现原理', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    tb(s, Emu(350000), Emu(1500000), Emu(11500000), Emu(350000),
       'sanitize_script(path) — 修复 Unicode 弯引号，然后用 ast.parse 验证语法',
       sz=12, c=C['navy'], C=C)
    code(s, Emu(350000), Emu(1950000), Emu(11500000), Emu(4400000),
         ['def sanitize_script(path):',
          '    """修复 Unicode 弯引号，然后做语法校验。"""',
          '    import ast',
          '    with open(path, encoding="utf-8") as f:',
          '        content = f.read()',
          '',
          '    # 替换 4 种 Unicode 弯引号',
          '    QUOTES = {',
          '        chr(0x201c): chr(34),   # U+201C 左双引号  ->  "',
          '        chr(0x201d): chr(34),   # U+201D 右双引号  ->  "',
          '        chr(0x2018): chr(39),   # U+2018 左单引号  ->  \'',
          '        chr(0x2019): chr(39),   # U+2019 右单引号  ->  \'',
          '    }',
          '    for old, new in QUOTES.items():',
          '        content = content.replace(old, new)',
          '',
          '    try:',
          '        ast.parse(content)',
          '        return True    # 语法正常，可以执行',
          '    except SyntaxError as e:',
          '        print(f"SyntaxError: {e}")',
          '        return False'],
         sz=10, C=C)
    sn(s, 9, TOTAL, C=C)


def build_slide_10_api_ref(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '核心 API 速查', 'pptx_helpers — 常用参数', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    hdr = ['函数', '关键参数', '调用示例']
    data = [
        ['create_prs(scheme)', 'navy_teal | tech_dark | corporate',
         "prs, C = create_prs('corporate')"],
        ['add_bg(s,l,t,w,h,c)', 'c: RGBColor',
         "add_bg(s, 0, 0, SLIDE_W, SLIDE_H, C['navy'])"],
        ['tb(s,l,t,w,h,txt)', 'sz, b, c, al, fn, ls',
         "tb(s, l, t, w, h, '标题', sz=18, b=True, c=C['navy'], C=C)"],
        ['ml(s,l,t,w,h,lines)', 'bullet, sz, ls',
         "ml(s, l, t, w, h, ['行1','行2'], bullet=True, C=C)"],
        ['ct(s,l,t,w,h,r,c,cw)', 'cw: 列宽列表（EMU）',
         'tbl = ct(s, l, t, w, h, 5, 3)'],
        ['ft(tbl, hdr, data)', 'fs, hfs, hc, hfc',
         "ft(tbl, ['列1','列2'], [['A','B']], C=C)"],
        ['code(s,l,t,w,h,lines)', 'sz',
         "code(s, l, t, w, h, ['import x', 'x()'], C=C)"],
        ['hbox(s,l,t,w,h,title,lines)', 'bg, tc',
         "hbox(s, l, t, w, h, '说明', ['要点'], bg=C['light_blue'], C=C)"],
        ['bar(s, title, sub)', 'C',
         "bar(s, '标题', sub='副标题', C=C)"],
        ['sn(s, n, total)', 'C',
         'sn(s, 1, 20, C=C)'],
    ]
    cw = [Emu(1800000), Emu(3800000), Emu(5000000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(10600000), Emu(5200000), 11, 3, cw)
    ft(t, hdr, data, fs=9, hfs=10, C=C)
    sn(s, 10, TOTAL, C=C)


def build_slide_11_color_schemes(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '配色方案对比', '3 套内置配色方案', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    hdr = ['方案键名', '调用方式', '风格', '主色', '适用场景']
    data = [
        ['navy_teal', "create_prs('navy_teal')", '深色', '#1B3A5C + #1A8A8A', '技术汇报、产品分析'],
        ['tech_dark', "create_prs('tech_dark')", '科技深色', '#0F0F23 + #00D9FF', 'AI/科技演示'],
        ['corporate', "create_prs('corporate')", '企业浅色', '#2E5C9A + 白色背景', '企业正式汇报（推荐）'],
    ]
    cw = [Emu(1200000), Emu(2200000), Emu(1200000), Emu(2000000), Emu(2700000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(9300000), Emu(2200000), 4, 5, cw)
    ft(t, hdr, data, fs=10, hfs=11, C=C)

    tb(s, Emu(350000), Emu(4000000), Emu(11500000), Emu(350000),
       '每套方案包含 21 个语义化颜色键', sz=13, b=True, c=C['navy'], C=C)
    ml(s, Emu(350000), Emu(4450000), Emu(5600000), Emu(2000000),
       ['基础色: navy, navy_dark, teal, white, black',
        '语义色: amber, red, green, orange',
        '灰度: gray_light, gray_mid, gray_text'],
       sz=11, bullet=True, C=C)
    ml(s, Emu(6050000), Emu(4450000), Emu(5600000), Emu(2000000),
       ['行色: row_alt, light_blue, light_red, light_green, light_amber',
        '代码色: code_bg, code_text',
        '字体: font_cn (微软雅黑), font_code (Courier New)'],
       sz=11, bullet=True, C=C)
    sn(s, 11, TOTAL, C=C)


def build_slide_12_builder_pattern(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, 'Builder 模式脚本骨架', '推荐的脚本结构', C=C)
    act_badge(s, Emu(350000), Emu(870000), '三', '技术实现', C['navy'], C=C)

    code(s, Emu(350000), Emu(1500000), Emu(11500000), Emu(5200000),
         ['from scripts.pptx_helpers import *',
          '',
          'SLIDE_PLAN = [',
          '    (1, "封面",   "全屏标题", "主题 + 副标题"),',
          '    (2, "目录",   "目录",     "章节概览"),',
          '    # ... 更多页面',
          ']',
          '',
          'prs, C = create_prs("corporate")',
          '',
          '',
          'def build_slide_01_cover(prs, C):',
          '    s = prs.slides.add_slide(prs.slide_layouts[6])',
          '    bar(s, "标题", sub="副标题", C=C)',
          '    sn(s, 1, len(SLIDE_PLAN), C=C)',
          '',
          '',
          'def main():',
          '    builder_batches = [',
          '        ("Batch 1", [build_slide_01_cover, ...]),',
          '        ("Batch 2", [build_slide_09_xxxx,  ...]),',
          '    ]',
          '    for batch_name, builders in builder_batches:',
          '        for b in builders:',
          '            b(prs, C)',
          '    sanitize_script(__file__)    # 代码卫生校验',
          '    save_pptx(prs, "output.pptx")'],
         sz=10, C=C)
    sn(s, 12, TOTAL, C=C)


def build_slide_13_comparison(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '改进前后对比', '关键指标变化', C=C)
    act_badge(s, Emu(350000), Emu(870000), '四', '改进成果', C['green'], C=C)

    hdr = ['指标', '改进前', '改进后', '变化']
    data = [
        ['脚本行数（典型 20 页）',   '~900 行',       '~500 行',       '↓ 44%'],
        ['Helper 定义代码',          '~150 行/脚本',   '1 行 import',   '↓ 99%'],
        ['多脚本重复代码',           '~600 行',        '0 行',          '彻底消除'],
        ['25 页 deck 成功率',        '~25%',           '待实测',        '预期提升'],
        ['SyntaxError 风险',         '需手动修复',     '自动检测+修复', '消除'],
        ['生成失败后处理',           '全部重来',       '断点续写',      '成本↓'],
    ]
    cw = [Emu(3500000), Emu(2200000), Emu(2200000), Emu(2100000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(10000000), Emu(4000000), 7, 4, cw)
    ft(t, hdr, data, fs=11, hfs=12, C=C)

    change_colors = [C['green'], C['green'], C['green'], C['teal'], C['green'], C['green']]
    change_texts = ['↓ 44%', '↓ 99%', '彻底消除', '预期提升', '消除', '成本↓']
    for row in range(1, 7):
        sc(t.cell(row, 3), change_texts[row - 1], sz=11,
           fc=change_colors[row - 1], C=C)
    sn(s, 13, TOTAL, C=C)


def build_slide_14_metrics(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '成果数据', '脚本体积与效率改善', C=C)
    act_badge(s, Emu(350000), Emu(870000), '四', '改进成果', C['green'], C=C)

    metrics = [
        ('↓ 44%', '脚本行数减少',  '900 行\n→ 500 行',    C['green']),
        ('↓ 99%', 'Helper 重复',   '150 行\n→ 1 行 import', C['teal']),
        ('0 行',  '重复配色定义',  '每次都写\n→ 完全复用', C['navy']),
        ('✓',     '代码卫生保证',  'Unicode+语法\n自动检测', C['amber']),
    ]
    for i, (value, label, detail, clr) in enumerate(metrics):
        x = Emu(500000 + i * 2850000)
        add_bg(s, x, Emu(1300000), Emu(2600000), Emu(3800000), clr)
        tb(s, x + Emu(100000), Emu(1450000), Emu(2400000), Emu(900000),
           value, sz=40, b=True, c=C['white'], al=PP_ALIGN.CENTER, C=C)
        tb(s, x + Emu(100000), Emu(2400000), Emu(2400000), Emu(380000),
           label, sz=13, b=True, c=C['white'], al=PP_ALIGN.CENTER, C=C)
        ml(s, x + Emu(100000), Emu(2850000), Emu(2400000), Emu(1100000),
           detail.split('\n'), sz=11, c=C['white'], al=PP_ALIGN.CENTER, C=C)

    hbox(s, Emu(350000), Emu(5400000), Emu(11500000), Emu(1000000),
         '注',
         ['脚本行数数据基于典型 20 页 deck；25 页以上 deck 的实测成功率待验证'],
         bg=C['light_amber'], tc=C['amber'], C=C)
    sn(s, 14, TOTAL, C=C)


def build_slide_15_conclusions(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '关键结论', C=C)
    act_badge(s, Emu(350000), Emu(870000), '五', '结论与展望', C['amber'], C=C)

    items = [
        ('✅', '已完成', 'pptx_helpers.py 落地',    'Helper 复用 + 配色方案 + sanitize_script，均已可用',          C['green']),
        ('✅', '已完成', 'agent 工作流更新',         '规划先行 → Helper 优先 → 分块生成 → 代码卫生 → 验证',          C['green']),
        ('✅', '已完成', 'builder 风格示例',         'example_helpers.py 已改造，运行正常并输出 PPTX',               C['green']),
        ('✅', '已完成', '16 页 deck 首次验证',      '当前脚本即为首次系统性验证（本次运行）',                        C['teal']),
        ('⏳', '待完成', '追加模式 + 断点续写演示',  '已写入 agent 工作流，但缺少现成示例脚本',                       C['amber']),
    ]
    for i, (icon, status, title, desc, clr) in enumerate(items):
        y = Emu(1500000 + i * 950000)
        badge(s, Emu(350000), y, Emu(480000), Emu(390000), icon, clr, C=C)
        tb(s, Emu(1000000), y + Emu(15000), Emu(1500000), Emu(360000),
           status, sz=11, b=True, c=clr, fn=C['font_code'], C=C)
        tb(s, Emu(2650000), y + Emu(15000), Emu(2600000), Emu(360000),
           title, sz=13, b=True, c=C['black'], C=C)
        tb(s, Emu(5400000), y + Emu(35000), Emu(6400000), Emu(330000),
           desc, sz=11, c=C['gray_text'], C=C)
    sn(s, 15, TOTAL, C=C)


def build_slide_16_closing(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, 0, 0, SLIDE_W, SLIDE_H, C['navy_dark'])
    add_bg(s, 0, Emu(2500000), SLIDE_W, Emu(55000), C['teal'])
    add_bg(s, 0, Emu(4400000), SLIDE_W, Emu(55000), C['teal'])
    tb(s, Emu(800000), Emu(1700000), Emu(10500000), Emu(600000),
       '后续计划', sz=26, b=True, c=C['teal'], al=PP_ALIGN.CENTER, C=C)
    ml(s, Emu(1500000), Emu(2600000), Emu(9200000), Emu(1700000),
       ['用 15~30 页真实业务 deck 进一步验证分块生成稳定性',
        '补充断点续写演示示例：从中断的 slide builder 列表恢复追加',
        '如验证通过，将摘要格式和 slide plan 表格固化为可复用模板'],
       sz=14, c=C['white'], bullet=True, ls=1.6, C=C)
    tb(s, Emu(800000), Emu(4600000), Emu(10500000), Emu(700000),
       'Q&A', sz=44, b=True, c=C['white'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(5400000), Emu(10500000), Emu(400000),
       'github.com/Ami-zhang/pptx-creator-agent',
       sz=12, c=C['teal'], al=PP_ALIGN.CENTER, fn=C['font_code'], C=C)
    sn(s, 16, TOTAL, C=C)


# ═══════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════

def main():
    prs, C = create_prs('navy_teal')
    print_slide_plan()

    builder_batches = [
        ('Batch 1 (slides 01–08)', [
            build_slide_01_cover,
            build_slide_02_agenda,
            build_slide_03_status,
            build_slide_04_problems,
            build_slide_05_fatal_issue,
            build_slide_06_solutions,
            build_slide_07_helpers_design,
            build_slide_08_chunking,
        ]),
        ('Batch 2 (slides 09–16)', [
            build_slide_09_hygiene,
            build_slide_10_api_ref,
            build_slide_11_color_schemes,
            build_slide_12_builder_pattern,
            build_slide_13_comparison,
            build_slide_14_metrics,
            build_slide_15_conclusions,
            build_slide_16_closing,
        ]),
    ]

    for batch_name, builders in builder_batches:
        run_builder_batch(prs, C, batch_name, builders)

    # 代码卫生校验（演示 sanitize_script 工作流）
    this_file = os.path.abspath(__file__)
    print(f'\nRunning sanitize_script on {os.path.basename(this_file)}...')
    ok = sanitize_script(this_file)
    print(f'  Syntax: {"OK" if ok else "FAILED — fix before proceeding"}')

    out = os.path.join(os.path.dirname(this_file), 'validate-15page-deck.pptx')
    save_pptx(prs, out)


if __name__ == '__main__':
    main()
