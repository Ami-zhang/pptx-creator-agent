# -*- coding: utf-8 -*-
"""
pptx_helpers 示例 — 使用 builder 风格生成 8 页演示文稿

演示 pptx_helpers.py 的所有核心 API:
  add_bg, tb, ml, rl, sc, ct, ft, code, hbox, badge, bar, act_badge, sn

同时演示推荐工作流:
   1. 先定义 slide plan
   2. 再编写具名 slide builders
   3. 按批次执行 builders

用法:
    python example_helpers.py
"""

import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))

from pptx_helpers import *

SLIDE_PLAN = [
   (1, '封面', '全屏标题', '介绍 helper 库定位与核心卖点'),
   (2, '目录', '目录页', '概览 6 个示例主题'),
   (3, '基础元素演示', '混合布局', '展示 add_bg / tb / ml'),
   (4, '表格演示', '表格 + 说明框', '展示 ct / ft / sc'),
   (5, '代码块与信息框', '双栏对比', '展示 code / hbox'),
   (6, '徽章与章节标记', '组件展示', '展示 badge / act_badge'),
   (7, '富文本列表与配色', '列表 + 表格 + 代码', '展示 rl / create_prs'),
   (8, '结束页', '全屏收束', '总结 helper 库适用场景'),
]

TOTAL = len(SLIDE_PLAN)


def print_slide_plan():
   """打印规划表，演示推荐的 planning-first 工作流。"""
   print('Slide plan:')
   for page, title, layout, content in SLIDE_PLAN:
      print(f'  {page:02d}. {title} | {layout} | {content}')


def run_builder_batch(prs, C, batch_name, builders):
   """按批次执行 slide builders。"""
   print(f'Running {batch_name} with {len(builders)} slides...')
   for builder in builders:
      builder(prs, C)


# ──────── S1: 封面 ────────
def build_slide_01_cover(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, 0, 0, SLIDE_W, SLIDE_H, C['navy_dark'])
    add_bg(s, 0, Emu(2400000), SLIDE_W, Emu(2600000), C['navy'])
    add_bg(s, 0, Emu(2400000), SLIDE_W, Emu(50000), C['teal'])
    add_bg(s, 0, Emu(4950000), SLIDE_W, Emu(50000), C['teal'])
    tb(s, Emu(800000), Emu(2700000), Emu(10500000), Emu(700000),
       'pptx_helpers \u793a\u4f8b\u6f14\u793a', sz=36, b=True, c=C['white'],
       al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(3400000), Emu(10500000), Emu(500000),
       '\u5143\u7d20\u7ea7\u7ec4\u4ef6\u5e93  |  3 \u5957\u914d\u8272  |  12 \u4e2a Helper  |  \u5f00\u7bb1\u5373\u7528',
       sz=14, c=C['gray_mid'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(4000000), Emu(10500000), Emu(400000),
       'from pptx_helpers import *',
       sz=16, c=C['teal'], al=PP_ALIGN.CENTER, fn=C['font_code'], C=C)
    sn(s, 1, TOTAL, C)


# ──────── S2: 目录 ────────
def build_slide_02_agenda(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u76ee\u5f55', C=C)
    items = [
        ('1', '\u57fa\u7840\u5143\u7d20', '\u77e9\u5f62\u8272\u5757 + \u6587\u672c\u6846 + \u591a\u884c\u6587\u672c', C['teal']),
        ('2', '\u8868\u683c\u4e0e\u5355\u5143\u683c', '\u521b\u5efa\u8868\u683c + \u586b\u5145 + \u6837\u5f0f\u5316', C['navy']),
        ('3', '\u4ee3\u7801\u5757\u4e0e\u4fe1\u606f\u6846', '\u6df1\u8272\u4ee3\u7801\u5757 + \u9ad8\u4eae\u6846', C['amber']),
        ('4', '\u5fbd\u7ae0\u4e0e\u7ae0\u8282\u6807\u8bb0', '\u6807\u7b7e\u5fbd\u7ae0 + \u5e55\u6807\u9898\u6761', C['red']),
        ('5', '\u5bcc\u6587\u672c\u5217\u8868', '\u6bcf\u884c\u72ec\u7acb\u6837\u5f0f\u7684\u5217\u8868', C['green']),
        ('6', '\u914d\u8272\u65b9\u6848\u5bf9\u6bd4', '3 \u5957\u5185\u7f6e\u914d\u8272\u65b9\u6848', C['orange']),
    ]
    for i, (num, title, desc, clr) in enumerate(items):
        y = Emu(1000000 + i * 850000)
        badge(s, Emu(400000), y, Emu(500000), Emu(400000), num, clr, C=C)
        tb(s, Emu(1100000), y + Emu(30000), Emu(4500000), Emu(350000),
           title, sz=18, b=True, c=clr, C=C)
        tb(s, Emu(5800000), y + Emu(60000), Emu(5500000), Emu(300000),
           desc, sz=12, c=C['gray_text'], C=C)
    sn(s, 2, TOTAL, C)


# ──────── S3: 基础元素 ────────
def build_slide_03_basics(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u57fa\u7840\u5143\u7d20\u6f14\u793a', 'add_bg + tb + ml', C=C)
    act_badge(s, Emu(350000), Emu(880000), '\u4e00', '\u57fa\u7840\u5143\u7d20', C['teal'], C=C)

    # add_bg \u793a\u4f8b
    add_bg(s, Emu(350000), Emu(1500000), Emu(3500000), Emu(2000000), C['light_blue'])
    tb(s, Emu(500000), Emu(1600000), Emu(3200000), Emu(350000),
       'add_bg() \u2014 \u77e9\u5f62\u8272\u5757', sz=14, b=True, c=C['navy'], C=C)
    ml(s, Emu(500000), Emu(2000000), Emu(3200000), Emu(1300000),
       ['\u7eaf\u8272\u586b\u5145\u77e9\u5f62',
        '\u65e0\u8fb9\u6846\u3001\u65e0\u9634\u5f71',
        '\u53ef\u4f5c\u4e3a\u80cc\u666f\u3001\u8272\u5757\u3001\u8868\u5934\u5e95\u8272',
        '\u6700\u5e38\u7528\u7684\u57fa\u7840\u5143\u7d20'],
       sz=12, c=C['black'], bullet=True, ls=1.4, C=C)

    # tb \u793a\u4f8b
    add_bg(s, Emu(4100000), Emu(1500000), Emu(3500000), Emu(2000000), C['light_green'])
    tb(s, Emu(4250000), Emu(1600000), Emu(3200000), Emu(350000),
       'tb() \u2014 \u5355\u884c\u6587\u672c\u6846', sz=14, b=True, c=C['green'], C=C)
    tb(s, Emu(4250000), Emu(2050000), Emu(3200000), Emu(300000),
       '\u5b57\u53f7\u3001\u52a0\u7c97\u3001\u989c\u8272\u3001\u5bf9\u9f50\u3001\u5b57\u4f53\u3001\u884c\u8ddd\u5168\u53ef\u63a7',
       sz=12, c=C['black'], C=C)
    tb(s, Emu(4250000), Emu(2450000), Emu(3200000), Emu(300000),
       '\u2190 24pt \u7c97\u4f53 Navy \u5c45\u4e2d', sz=11, c=C['gray_text'], C=C)
    tb(s, Emu(4250000), Emu(2800000), Emu(3200000), Emu(300000),
       '\u2190 12pt \u666e\u901a Green \u5de6\u5bf9\u9f50', sz=11, c=C['gray_text'], C=C)

    # ml \u793a\u4f8b
    add_bg(s, Emu(7850000), Emu(1500000), Emu(3900000), Emu(2000000), C['light_amber'])
    tb(s, Emu(8000000), Emu(1600000), Emu(3600000), Emu(350000),
       'ml() \u2014 \u591a\u884c\u6587\u672c', sz=14, b=True, c=C['amber'], C=C)
    ml(s, Emu(8000000), Emu(2000000), Emu(3600000), Emu(1300000),
       ['\u652f\u6301 bullet \u9879\u76ee\u7b26\u53f7',
        '\u884c\u8ddd\u53ef\u8c03 (ls=1.3)',
        '\u53ef\u8bbe\u7f6e\u52a0\u7c97\u3001\u5bf9\u9f50',
        '\u9002\u5408\u5217\u4e3e\u8981\u70b9'],
       sz=12, c=C['black'], bullet=True, ls=1.4, C=C)

    # \u4ee3\u7801\u5c55\u793a
    tb(s, Emu(400000), Emu(3800000), Emu(11000000), Emu(350000),
       '\u8c03\u7528\u793a\u4f8b', sz=15, b=True, c=C['navy'], C=C)
    code(s, Emu(350000), Emu(4200000), Emu(11300000), Emu(2100000),
         ["add_bg(s, Emu(350000), Emu(1500000), Emu(3500000), Emu(2000000), C['light_blue'])",
          "",
          "tb(s, l, t, w, h, '\u6807\u9898\u6587\u672c', sz=24, b=True, c=C['navy'], C=C)",
          "",
          "ml(s, l, t, w, h,",
          "   ['\u7b2c\u4e00\u884c', '\u7b2c\u4e8c\u884c', '\u7b2c\u4e09\u884c'],",
          "   sz=12, bullet=True, ls=1.4, C=C)"],
         sz=10, C=C)
    sn(s, 3, TOTAL, C)


# ──────── S4: 表格 ────────
def build_slide_04_tables(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u8868\u683c\u6f14\u793a', 'ct + ft + sc', C=C)
    act_badge(s, Emu(350000), Emu(880000), '\u4e8c', '\u8868\u683c\u4e0e\u5355\u5143\u683c', C['navy'], C=C)

    hdr = ['Helper', '\u7528\u9014', '\u4f7f\u7528\u9891\u6b21', '\u7c92\u5ea6']
    data = [
        ['add_bg()', '\u7eaf\u8272\u77e9\u5f62\u80cc\u666f', '\u6781\u9ad8', '\u5143\u7d20\u7ea7'],
        ['tb()', '\u5355\u884c\u6587\u672c\u6846', '\u6781\u9ad8', '\u5143\u7d20\u7ea7'],
        ['ml()', '\u591a\u884c\u6587\u672c + bullet', '\u9ad8', '\u5143\u7d20\u7ea7'],
        ['ct() + ft()', '\u521b\u5efa\u8868\u683c + \u586b\u5145\u6570\u636e', '\u9ad8', '\u5143\u7d20\u7ea7'],
        ['sc()', '\u8bbe\u7f6e\u5355\u5143\u683c\u6837\u5f0f', '\u9ad8', '\u5355\u5143\u683c\u7ea7'],
        ['code()', '\u6df1\u8272\u4ee3\u7801\u5757', '\u4e2d', '\u5143\u7d20\u7ea7'],
        ['hbox()', '\u5e26\u6807\u9898\u7684\u4fe1\u606f\u6846', '\u9ad8', '\u5143\u7d20\u7ea7'],
        ['badge()', '\u5e8f\u53f7/\u6807\u7b7e\u5fbd\u7ae0', '\u4e2d', '\u5143\u7d20\u7ea7'],
    ]
    cw = [Emu(2000000), Emu(3200000), Emu(1500000), Emu(1500000)]
    t = ct(s, Emu(350000), Emu(1500000), Emu(8200000), Emu(4500000), 9, 4, cw)
    ft(t, hdr, data, fs=11, hfs=12, C=C)

    # \u53f3\u4fa7\uff1a\u5355\u5143\u683c\u6837\u5f0f\u5316\u793a\u4f8b
    add_bg(s, Emu(8800000), Emu(1500000), Emu(3100000), Emu(4500000), C['light_blue'])
    tb(s, Emu(8950000), Emu(1600000), Emu(2800000), Emu(350000),
       'sc() \u5355\u5143\u683c\u6837\u5f0f\u5316', sz=13, b=True, c=C['navy'], C=C)
    code(s, Emu(8900000), Emu(2050000), Emu(3000000), Emu(3700000),
         ["# \u8bbe\u7f6e\u5355\u5143\u683c\u6837\u5f0f",
          "sc(tbl.cell(1,2),",
          "   '\u2705 \u901a\u8fc7',",
          "   sz=12,",
          "   b=True,",
          "   fc=C['green'],",
          "   bg=C['light_green'],",
          "   al=PP_ALIGN.CENTER,",
          "   C=C)",
          "",
          "# \u7ea2\u8272\u8b66\u544a\u5355\u5143\u683c",
          "sc(tbl.cell(2,2),",
          "   '\u274c \u5931\u8d25',",
          "   fc=C['red'],",
          "   bg=C['light_red'],",
          "   C=C)"],
         sz=9, C=C)
    sn(s, 4, TOTAL, C)


# ──────── S5: 代码块与信息框 ────────
def build_slide_05_code_and_boxes(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u4ee3\u7801\u5757\u4e0e\u4fe1\u606f\u6846', 'code + hbox', C=C)
    act_badge(s, Emu(350000), Emu(880000), '\u4e09', '\u4ee3\u7801\u5757\u4e0e\u4fe1\u606f\u6846', C['amber'], C=C)

    # code \u793a\u4f8b
    tb(s, Emu(400000), Emu(1500000), Emu(5000000), Emu(350000),
       'code() \u2014 \u6df1\u8272\u80cc\u666f\u4ee3\u7801\u5757', sz=14, b=True, c=C['navy'], C=C)
    code(s, Emu(350000), Emu(1900000), Emu(5500000), Emu(1800000),
         ['from pptx_helpers import *',
          '',
          'prs, C = create_prs("navy_teal")',
          '',
          'def build_slide_01(prs, C):',
          '    s = prs.slides.add_slide(prs.slide_layouts[6])',
          '    bar(s, "\u6807\u9898", "\u526f\u6807\u9898", C=C)',
          '    sn(s, 1, 10, C=C)'],
         sz=11, C=C)

    # hbox \u793a\u4f8b
    tb(s, Emu(6200000), Emu(1500000), Emu(5500000), Emu(350000),
       'hbox() \u2014 \u5e26\u6807\u9898\u7684\u9ad8\u4eae\u4fe1\u606f\u6846', sz=14, b=True, c=C['navy'], C=C)
    hbox(s, Emu(6100000), Emu(1900000), Emu(5600000), Emu(1800000),
         '\u4f18\u70b9',
         ['\u652f\u6301\u6807\u9898 + \u591a\u884c\u5185\u5bb9',
          '\u53ef\u81ea\u5b9a\u4e49\u80cc\u666f\u8272\u548c\u6807\u9898\u8272',
          '\u81ea\u5e26 bullet \u5217\u8868',
          '\u9002\u5408\u5de6\u53f3\u5bf9\u6bd4\u5e03\u5c40'],
         bg=C['light_green'], tc=C['green'], C=C)

    # \u4e24\u4e2a hbox \u5bf9\u6bd4\u5e03\u5c40\u793a\u4f8b
    tb(s, Emu(400000), Emu(4000000), Emu(11000000), Emu(350000),
       '\u53cc\u680f\u5bf9\u6bd4\u5e03\u5c40\u793a\u4f8b', sz=15, b=True, c=C['navy'], C=C)
    hbox(s, Emu(350000), Emu(4400000), Emu(5500000), Emu(2000000),
         '\u65b9\u6848 A\uff1a\u77ed\u671f\u4fee\u590d',
         ['\u96f6\u6210\u672c\u5b9e\u65bd',
          '\u4e0d\u9700\u8981\u91cd\u65b0\u7f16\u8bd1',
          '\u8986\u76d6 80% \u573a\u666f'],
         bg=C['light_blue'], tc=C['navy'], C=C)
    hbox(s, Emu(6100000), Emu(4400000), Emu(5600000), Emu(2000000),
         '\u65b9\u6848 B\uff1a\u67b6\u6784\u5347\u7ea7',
         ['\u9700\u8981 1-3 \u4e2a\u6708\u5f00\u53d1',
          '\u6839\u672c\u89e3\u51b3\u95ee\u9898',
          '\u8986\u76d6 100% \u573a\u666f'],
         bg=C['light_amber'], tc=C['amber'], C=C)
    sn(s, 5, TOTAL, C)


# ──────── S6: 徽章与章节标记 ────────
def build_slide_06_badges(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u5fbd\u7ae0\u4e0e\u7ae0\u8282\u6807\u8bb0', 'badge + act_badge', C=C)
    act_badge(s, Emu(350000), Emu(880000), '\u56db', '\u5fbd\u7ae0\u4e0e\u7ae0\u8282\u6807\u8bb0', C['red'], C=C)

    # badge \u793a\u4f8b
    tb(s, Emu(400000), Emu(1500000), Emu(5000000), Emu(350000),
       'badge() \u2014 \u5e8f\u53f7/\u6807\u7b7e\u5fbd\u7ae0', sz=14, b=True, c=C['navy'], C=C)
    colors = [C['navy'], C['teal'], C['amber'], C['red'], C['green'], C['orange']]
    labels = ['1', '2', '3', 'P0', 'OK', '\u2605']
    for i, (clr, lbl) in enumerate(zip(colors, labels)):
        x = Emu(400000 + i * 1000000)
        badge(s, x, Emu(1950000), Emu(700000), Emu(400000), lbl, clr, C=C)

    # act_badge \u793a\u4f8b
    tb(s, Emu(400000), Emu(2700000), Emu(11000000), Emu(350000),
       'act_badge() \u2014 \u7ae0\u8282\u6807\u8bb0\u6761\uff08\u6bcf\u4e00\u5e55\u7684\u5f00\u5934\uff09', sz=14, b=True, c=C['navy'], C=C)
    act_badge(s, Emu(350000), Emu(3200000), '\u4e00', '\u95ee\u9898\u53d1\u73b0', C['red'], C=C)
    act_badge(s, Emu(350000), Emu(3750000), '\u4e8c', '\u7cfb\u7edf\u6d4b\u8bd5', C['teal'], C=C)
    act_badge(s, Emu(350000), Emu(4300000), '\u4e09', '\u89e3\u51b3\u65b9\u6848', C['green'], C=C)
    act_badge(s, Emu(350000), Emu(4850000), '\u56db', '\u7ade\u54c1\u5bf9\u6bd4', C['amber'], C=C)

    # \u4ee3\u7801
    code(s, Emu(350000), Emu(5500000), Emu(11300000), Emu(800000),
         ["badge(s, l, t, w, h, 'P0', C['red'], C=C)",
          "",
          "act_badge(s, Emu(350000), Emu(880000), '\u4e00', '\u7ae0\u8282\u6807\u9898', C['teal'], C=C)"],
         sz=11, C=C)
    sn(s, 6, TOTAL, C)


# ──────── S7: 富文本列表 + 配色对比 ────────
def build_slide_07_richtext_and_schemes(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar(s, '\u5bcc\u6587\u672c\u5217\u8868 & \u914d\u8272\u65b9\u6848', 'rl + create_prs', C=C)

    # rl \u793a\u4f8b
    tb(s, Emu(400000), Emu(950000), Emu(5000000), Emu(350000),
       'rl() \u2014 \u6bcf\u884c\u72ec\u7acb\u6837\u5f0f', sz=14, b=True, c=C['navy'], C=C)
    items = [
        ('\u2605 P0\uff1a\u7acb\u5373\u6267\u884c\u7684\u4fee\u590d', 14, True, C['red']),
        ('\u2605 P1\uff1a\u77ed\u671f\u4f18\u5316\u65b9\u6848', 14, True, C['amber']),
        ('\u2605 P2\uff1a\u4e2d\u671f\u67b6\u6784\u5347\u7ea7', 14, True, C['teal']),
        ('\u2014 \u666e\u901a\u6587\u672c\u884c', 12, False, C['black']),
        ('\u2014 \u7070\u8272\u6ce8\u91ca\u884c', 11, False, C['gray_text']),
    ]
    rl(s, Emu(400000), Emu(1400000), Emu(5000000), Emu(2200000), items, ls=1.5, C=C)

    # 3 \u5957\u914d\u8272\u65b9\u6848
    tb(s, Emu(6200000), Emu(950000), Emu(5500000), Emu(350000),
       '3 \u5957\u5185\u7f6e\u914d\u8272\u65b9\u6848', sz=14, b=True, c=C['navy'], C=C)
    hdr = ['\u65b9\u6848', '\u4e3b\u8272', '\u8f85\u8272', '\u9002\u7528\u573a\u666f']
    data = [
        ['navy_teal', 'Navy #1B3A5C', 'Teal #1A8A8A', '\u6280\u672f\u6c47\u62a5\u3001\u4ea7\u54c1\u5206\u6790'],
        ['tech_dark', 'Dark #0F0F23', 'Cyan #00D9FF', 'AI/\u79d1\u6280\u3001\u6df1\u8272\u6f14\u793a'],
        ['corporate', 'Blue #2E5C9A', 'Blue #2E5C9A', '\u4f01\u4e1a\u6b63\u5f0f\u3001\u767d\u5e95\u84dd\u5934'],
    ]
    cw = [Emu(1200000), Emu(1500000), Emu(1500000), Emu(1800000)]
    t = ct(s, Emu(6100000), Emu(1400000), Emu(6000000), Emu(2000000), 4, 4, cw)
    ft(t, hdr, data, fs=10, hfs=10, C=C)

    # \u4ee3\u7801\u793a\u4f8b
    code(s, Emu(350000), Emu(4000000), Emu(11300000), Emu(2400000),
         ["# \u5bcc\u6587\u672c\u5217\u8868",
          "items = [",
          "    ('P0 \u7d27\u6025\u4fee\u590d', 14, True, C['red']),",
          "    ('P1 \u4f18\u5316\u65b9\u6848', 14, True, C['amber']),",
          "    ('\u666e\u901a\u8bf4\u660e',    12, False, C['black']),",
          "]",
          "rl(s, l, t, w, h, items, ls=1.5, C=C)",
          "",
        "# \u6279\u6b21\u6267\u884c builders",
        "builder_batches = [",
        "    [build_slide_01, build_slide_02],",
        "]",
        "prs, C = create_prs('corporate')   # \u4f01\u4e1a\u6b63\u5f0f\u98ce"],
         sz=10, C=C)
    sn(s, 7, TOTAL, C)


# ──────── S8: 结束页 ────────
def build_slide_08_closing(prs, C):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, 0, 0, SLIDE_W, SLIDE_H, C['navy_dark'])
    add_bg(s, 0, Emu(2600000), SLIDE_W, Emu(50000), C['teal'])
    add_bg(s, 0, Emu(4400000), SLIDE_W, Emu(50000), C['teal'])
    tb(s, Emu(800000), Emu(2900000), Emu(10500000), Emu(600000),
       'pptx_helpers', sz=36, b=True, c=C['white'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(3500000), Emu(10500000), Emu(500000),
       '\u5143\u7d20\u7ea7\u7ec4\u4ef6\u5e93  \u00b7  \u5f00\u7bb1\u5373\u7528  \u00b7  \u4e13\u4e3a\u590d\u6742\u6df7\u5408\u5e03\u5c40\u8bbe\u8ba1',
       sz=14, c=C['gray_mid'], al=PP_ALIGN.CENTER, C=C)
    tb(s, Emu(800000), Emu(4700000), Emu(10500000), Emu(400000),
       'github.com/Ami-zhang/pptx-creator-agent',
       sz=12, c=C['teal'], al=PP_ALIGN.CENTER, fn=C['font_code'], C=C)
    sn(s, 8, TOTAL, C)


# ═══════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════
def main():
    prs, C = create_prs('navy_teal')
    print_slide_plan()

    builder_batches = [
        ('Batch 1', [
            build_slide_01_cover,
            build_slide_02_agenda,
            build_slide_03_basics,
            build_slide_04_tables,
        ]),
        ('Batch 2', [
            build_slide_05_code_and_boxes,
            build_slide_06_badges,
            build_slide_07_richtext_and_schemes,
            build_slide_08_closing,
        ]),
    ]

    for batch_name, builders in builder_batches:
        run_builder_batch(prs, C, batch_name, builders)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       'example-helpers.pptx')
    save_pptx(prs, out)


if __name__ == '__main__':
    main()
