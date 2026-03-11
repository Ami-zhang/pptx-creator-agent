# -*- coding: utf-8 -*-
"""
PPTX Creator - 通用演示文稿生成模板
使用 python-pptx 生成专业的 PowerPoint 演示文稿

使用方法:
    python pptx_template.py

依赖:
    pip install python-pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os


# ============================================================
# 配色方案 - 可根据主题修改
# ============================================================

# 科技主题配色
COLORS_TECH = {
    'bg_dark': RGBColor(0x0F, 0x0F, 0x23),
    'accent_blue': RGBColor(0x66, 0x7E, 0xEA),
    'accent_purple': RGBColor(0x76, 0x4B, 0xA2),
    'accent_cyan': RGBColor(0x00, 0xD9, 0xFF),
    'white': RGBColor(0xFF, 0xFF, 0xFF),
    'gray': RGBColor(0xAA, 0xAA, 0xAA),
    'card_bg': RGBColor(0x1A, 0x1A, 0x3E),
    'line_gray': RGBColor(0x33, 0x33, 0x55),
}

# 商务主题配色
COLORS_BUSINESS = {
    'bg_dark': RGBColor(0x1C, 0x28, 0x33),
    'accent_blue': RGBColor(0x2E, 0x86, 0xAB),
    'accent_purple': RGBColor(0x2C, 0x3E, 0x50),
    'accent_cyan': RGBColor(0xF3, 0x9C, 0x12),
    'white': RGBColor(0xFF, 0xFF, 0xFF),
    'gray': RGBColor(0xBD, 0xC3, 0xC7),
    'card_bg': RGBColor(0x2E, 0x40, 0x53),
    'line_gray': RGBColor(0x34, 0x49, 0x5E),
}

# 选择配色方案
COLORS = COLORS_TECH


# ============================================================
# 工具函数
# ============================================================

def set_shape_fill(shape, color):
    """设置形状填充颜色"""
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def add_text_frame(shape, text, font_size, color, bold=False, alignment=PP_ALIGN.LEFT):
    """添加文本到形状"""
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = 'Arial'
    p.alignment = alignment
    return tf


def add_slide_background(slide, prs, color):
    """添加幻灯片背景"""
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    set_shape_fill(background, color)
    background.line.fill.background()
    return background


def add_header_bar(slide, prs, title, color):
    """添加标题栏"""
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.8)
    )
    set_shape_fill(header, color)
    header.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.18), Inches(9), Inches(0.5))
    add_text_frame(title_box, title, 28, COLORS['white'], bold=True)
    return header


# ============================================================
# 幻灯片创建函数
# ============================================================

def create_title_slide(prs, title, subtitle, tagline=""):
    """创建标题页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    add_slide_background(slide, prs, COLORS['bg_dark'])
    
    # 顶部强调线
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Pt(8))
    set_shape_fill(top_bar, COLORS['accent_blue'])
    top_bar.line.fill.background()
    
    # 装饰圆形
    circle1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.5), Inches(4), Inches(1.2), Inches(1.2))
    set_shape_fill(circle1, COLORS['accent_purple'])
    circle1.line.fill.background()
    
    # 主标题
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(9), Inches(1))
    add_text_frame(title_box, title, 48, COLORS['white'], bold=True, alignment=PP_ALIGN.CENTER)
    
    # 副标题
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9), Inches(0.6))
    add_text_frame(subtitle_box, subtitle, 28, COLORS['accent_blue'], alignment=PP_ALIGN.CENTER)
    
    # 标语
    if tagline:
        tagline_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(0.5))
        add_text_frame(tagline_box, tagline, 14, COLORS['gray'], alignment=PP_ALIGN.CENTER)
    
    return slide


def create_content_slide(prs, title, items):
    """创建内容列表页
    
    Args:
        prs: Presentation 对象
        title: 幻灯片标题
        items: 列表项 [(标题, 描述), ...]
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    add_slide_background(slide, prs, COLORS['bg_dark'])
    
    # 标题栏
    add_header_bar(slide, prs, title, COLORS['accent_blue'])
    
    # 内容项
    y_start = Inches(1.1)
    for i, (item_title, item_desc) in enumerate(items):
        y = y_start + Inches(i * 0.85)
        
        # 卡片背景
        card = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), y, Inches(9), Inches(0.65)
        )
        set_shape_fill(card, COLORS['card_bg'])
        card.line.fill.background()
        
        # 左边框
        left_border = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), y, Pt(4), Inches(0.65)
        )
        set_shape_fill(left_border, COLORS['accent_cyan'])
        left_border.line.fill.background()
        
        # 文本
        text_tb = slide.shapes.add_textbox(Inches(0.7), y + Inches(0.15), Inches(8.5), Inches(0.45))
        tf = text_tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run1 = p.add_run()
        run1.text = item_title
        run1.font.size = Pt(14)
        run1.font.color.rgb = COLORS['accent_cyan']
        run1.font.bold = True
        run1.font.name = 'Arial'
        if item_desc:
            run2 = p.add_run()
            run2.text = f" - {item_desc}"
            run2.font.size = Pt(14)
            run2.font.color.rgb = COLORS['white']
            run2.font.name = 'Arial'
    
    return slide


def create_cards_slide(prs, title, cards):
    """创建卡片网格页
    
    Args:
        prs: Presentation 对象
        title: 幻灯片标题
        cards: 卡片数据 [(标题, 描述), ...]  最多6个
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    add_slide_background(slide, prs, COLORS['bg_dark'])
    
    # 标题栏
    add_header_bar(slide, prs, title, COLORS['accent_blue'])
    
    # 卡片布局
    card_width = Inches(2.9)
    card_height = Inches(1.8)
    x_positions = [Inches(0.45), Inches(3.55), Inches(6.65)]
    y_positions = [Inches(1.0), Inches(3.0)]
    accent_colors = [COLORS['accent_blue'], COLORS['accent_purple'], COLORS['accent_cyan']]
    
    for i, (card_title, card_desc) in enumerate(cards[:6]):
        col = i % 3
        row = i // 3
        x = x_positions[col]
        y = y_positions[row]
        accent_color = accent_colors[i % 3]
        
        # 卡片背景
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_width, card_height)
        set_shape_fill(card, COLORS['card_bg'])
        card.line.fill.background()
        
        # 顶部边框
        top_border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_width, Pt(4))
        set_shape_fill(top_border, accent_color)
        top_border.line.fill.background()
        
        # 标题
        title_tb = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.2), card_width - Inches(0.3), Inches(0.4))
        add_text_frame(title_tb, card_title, 15, COLORS['white'], bold=True)
        
        # 描述
        desc_tb = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.55), card_width - Inches(0.3), Inches(1.2))
        tf = desc_tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = card_desc
        p.font.size = Pt(10)
        p.font.color.rgb = COLORS['gray']
        p.font.name = 'Arial'
    
    return slide


def create_summary_slide(prs, title, message):
    """创建总结页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    add_slide_background(slide, prs, COLORS['bg_dark'])
    
    # 标题栏
    add_header_bar(slide, prs, title, COLORS['accent_blue'])
    
    # 总结框
    summary_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(2.5), Inches(8), Inches(0.8)
    )
    set_shape_fill(summary_bar, COLORS['accent_blue'])
    summary_bar.line.fill.background()
    
    summary_text = slide.shapes.add_textbox(Inches(1), Inches(2.7), Inches(8), Inches(0.5))
    add_text_frame(summary_text, message, 18, COLORS['white'], bold=True, alignment=PP_ALIGN.CENTER)
    
    return slide


# ============================================================
# 主函数 - 示例用法
# ============================================================

def main():
    """示例：创建一个演示文稿"""
    
    # 创建演示文稿 (16:9)
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    print("创建演示文稿...")
    
    # 幻灯片 1: 标题页
    create_title_slide(
        prs,
        title="演示文稿标题",
        subtitle="副标题内容",
        tagline="这是一个使用 python-pptx 生成的演示文稿"
    )
    print("  ✓ 标题页")
    
    # 幻灯片 2: 内容列表
    create_content_slide(
        prs,
        title="主要内容",
        items=[
            ("第一点", "这是第一个要点的详细说明"),
            ("第二点", "这是第二个要点的详细说明"),
            ("第三点", "这是第三个要点的详细说明"),
            ("第四点", "这是第四个要点的详细说明"),
        ]
    )
    print("  ✓ 内容列表页")
    
    # 幻灯片 3: 卡片网格
    create_cards_slide(
        prs,
        title="核心要素",
        cards=[
            ("要素一", "这是第一个要素的详细描述内容，可以包含多行文字。"),
            ("要素二", "这是第二个要素的详细描述内容。"),
            ("要素三", "这是第三个要素的详细描述内容。"),
            ("要素四", "这是第四个要素的详细描述内容。"),
            ("要素五", "这是第五个要素的详细描述内容。"),
            ("要素六", "这是第六个要素的详细描述内容。"),
        ]
    )
    print("  ✓ 卡片网格页")
    
    # 幻灯片 4: 总结页
    create_summary_slide(
        prs,
        title="总结",
        message="感谢您的关注！"
    )
    print("  ✓ 总结页")
    
    # 保存
    output_path = os.path.join(os.path.dirname(__file__), "example-presentation.pptx")
    prs.save(output_path)
    
    print(f"\n✅ 演示文稿创建成功!")
    print(f"📁 输出文件: {output_path}")
    
    return output_path


if __name__ == "__main__":
    main()
