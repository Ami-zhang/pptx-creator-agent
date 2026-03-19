# PPTX Creator Agent

一个 VS Code Copilot 自定义 Agent，用于从文本内容自动生成专业的 PowerPoint 演示文稿。

## 功能特点

- 🎨 **智能设计** - 根据内容主题自动选择配色方案（科技/商务/企业浅色）
- 📝 **多种输入** - 支持直接文本、Markdown 文件、现有 PPTX 模板等
- 🎯 **专业模板** - 提供多种幻灯片布局（标题页、内容页、卡片页、表格页、总结页）
- 🧩 **元素级 Helper** - 12 个精细构建函数（背景块、文本框、表格、代码块、徽章等），支持复杂混合布局
- 📊 **表格支持** - 创建表格、合并单元格、自定义样式，满足企业 PPT 需求
- 🀄 **中文友好** - 内置微软雅黑/等线等 CJK 字体配置，告别 Arial 乱码
- 🔍 **模板分析** - 分析现有 PPTX 结构（shapes、表格布局、样式），辅助复刻
- 🛡️ **代码卫生** - 内置 `sanitize_script()` 自动检测并修复 Unicode 引号等常见语法问题
- 🔧 **易于扩展** - 模块化的 Python 代码，方便自定义

## 安装

### 1. 安装依赖

**方式 A: 在线安装（推荐）**

```bash
pip install python-pptx
```

**方式 B: 离线安装（Windows 64位 + Python 3.11）**

本项目提供了离线安装包，详见 [packages/README.md](packages/README.md)：

```bash
cd packages
pip install *.whl
```

### 2. 安装 Agent

**方式 A: 用户级安装（推荐）**

将 `.github/agents/pptx-creator.agent.md` 复制到：

| 操作系统 | 路径 |
|---------|------|
| Windows | `%APPDATA%\Code\User\prompts\` |
| macOS | `~/Library/Application Support/Code/User/prompts/` |
| Linux | `~/.config/Code/User/prompts/` |

**方式 B: 项目级安装**

将整个 `.github/agents/` 目录复制到你的项目根目录。

### 3. 重启 VS Code

## 使用方法

### 在 VS Code 中使用

1. 打开 Copilot Chat
2. 选择 `@pptx-creator` agent
3. 输入你的请求，例如：

```
创建一个关于人工智能发展的5页演示文稿
```

```
把以下内容做成PPT:
- 项目背景
- 技术方案
- 实施计划
- 预期成果
```

### 使用 Python 脚本

#### 方式 A: 整页级模板 (`pptx_template.py`)

```python
from scripts.pptx_template import (
    create_title_slide,
    create_content_slide,
    create_cards_slide,
    create_table_slide,
    create_summary_slide,
    analyze_template,
    COLORS_CORPORATE,
)
import scripts.pptx_template as tmpl
from pptx import Presentation
from pptx.util import Inches

# 切换为企业浅色主题（可选）
tmpl.COLORS = COLORS_CORPORATE

# 创建演示文稿 (标准 PowerPoint 16:9)
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 添加幻灯片
create_title_slide(prs, "标题", "副标题")
create_content_slide(prs, "内容", [("要点", "说明")])
create_table_slide(prs, "数据", ["列1", "列2"], [["A", "B"]])

# 保存
prs.save("output.pptx")
```

#### 方式 B: 元素级 Helper (`pptx_helpers.py`)

适用于需要精细控制布局的复杂演示文稿：

```python
from scripts.pptx_helpers import *
from pptx.util import Emu

# 创建 Presentation，选择配色方案
prs, C = create_prs('navy_teal')  # 或 'tech_dark', 'corporate'
s = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

# 顶部装饰条 + 标题
bar(s, '项目概览', sub='2025年度技术汇报', C=C)

# 文本框
tb(s, Emu(500000), Emu(1000000), Emu(5000000), Emu(400000),
   '关键发现', sz=18, b=True, c=C['navy'], C=C)

# 多行列表
ml(s, Emu(500000), Emu(1500000), Emu(5000000), Emu(2000000),
   ['发现一: 性能提升30%', '发现二: 覆盖率达95%', '发现三: 零关键缺陷'],
   bullet=True, C=C)

# 表格
tbl = ct(s, Emu(500000), Emu(3800000), Emu(11000000), Emu(2000000), 4, 3)
ft(tbl, ['指标', '目标', '实际'],
   [['性能', '>90%', '95%'], ['覆盖率', '>80%', '85%'], ['缺陷', '<5', '2']],
   C=C)

# 页码
sn(s, 1, 10, C=C)

# 保存
save_pptx(prs, 'output.pptx')
```

## 配色方案

### pptx_template.py 配色方案

| 方案 | 常量 | 风格 | 背景色 |
|------|------|------|--------|
| 科技主题 | `COLORS_TECH` | 深色 | #0F0F23 深蓝黑 |
| 商务主题 | `COLORS_BUSINESS` | 深色 | #1C2833 深灰蓝 |
| 企业主题 ⭐ | `COLORS_CORPORATE` | 浅色 | #FFFFFF + 蓝色表头 #2E5C9A |

可以在 `pptx_template.py` 中修改 `COLORS = COLORS_CORPORATE` 来切换。

### pptx_helpers.py 配色方案

| 方案 | 键名 | 风格 | 主色调 |
|------|------|------|--------|
| 海军蓝青 | `navy_teal` | 深色 | #1B3A5C 海军蓝 + #1A8A8A 青色 |
| 科技暗黑 | `tech_dark` | 深色 | #0F0F23 深蓝黑 + #00D9FF 电子蓝 |
| 企业经典 | `corporate` | 浅色 | #2E5C9A 蓝色 + 白色背景 |

通过 `create_prs('navy_teal')` 选择。每种方案包含 21 个语义化颜色键（`navy`, `teal`, `amber`, `red`, `green`, `code_bg` 等）和字体配置（`font_cn`, `font_code`）。

每种配色方案均包含字体配置，确保中英文混排正确渲染。

## 幻灯片类型

### 整页级模板 (`pptx_template.py`)

| 类型 | 函数 | 用途 |
|------|------|------|
| 标题页 | `create_title_slide()` | 演示文稿首页 |
| 内容页 | `create_content_slide()` | 列表形式的要点 |
| 卡片页 | `create_cards_slide()` | 网格布局，最多6个卡片 |
| 表格页 | `create_table_slide()` | 带样式表头的数据表格 |
| 总结页 | `create_summary_slide()` | 结束语、感谢页 |

#### 表格工具函数

| 函数 | 用途 |
|------|------|
| `create_table()` | 底层表格创建 |
| `set_cell()` | 设置单元格内容/样式（支持多行文本） |
| `merge_cells()` | 合并单元格（需先合并再写内容） |
| `set_table_style()` | 设置表格样式（通过 XML） |

#### 布局辅助

| 函数 | 用途 |
|------|------|
| `add_section_label()` | 分节标签（如蓝底白字横条） |
| `analyze_template()` | 分析现有 PPTX 模板结构 |

### 元素级 Helper 函数 (`pptx_helpers.py`)

从 4 个实战项目中提炼，与整页级模板互补使用，适合需要精细控制布局的复杂演示文稿。

| 函数 | 用途 | 参数概要 |
|------|------|----------|
| `create_prs()` | 工厂函数，创建 Presentation + 配色方案 | `scheme='navy_teal'` |
| `add_bg()` | 纯色矩形背景块 | `slide, l, t, w, h, color` |
| `tb()` | 单行/多行文本框 | `slide, l, t, w, h, text, sz, b, c, al, fn, ls` |
| `ml()` | 多行文本列表 | `slide, l, t, w, h, lines, bullet=True` |
| `rl()` | 富文本列表（每行独立样式） | `items=[(text, size, bold, color), ...]` |
| `ct()` | 创建表格 | `slide, l, t, w, h, rows, cols, cw` |
| `ft()` | 填充表格（表头+数据，自动交替行色） | `tbl, hdr, data` |
| `sc()` | 设置单元格内容和样式 | `cell, text, sz, b, fc, bg, al` |
| `code()` | 深色背景代码块 | `slide, l, t, w, h, lines` |
| `hbox()` | 带标题的高亮信息框 | `slide, l, t, w, h, title, lines` |
| `badge()` | 序号/标签徽章 | `slide, l, t, w, h, text, bg` |
| `act_badge()` | 章节标记条（幕标题） | `slide, l, t, act_num, act_title, color` |
| `bar()` | 顶部装饰条 + 标题 + 底部条 | `slide, title, sub` |
| `sn()` | 右下角页码标注 | `slide, n, total` |
| `save_pptx()` | 保存并打印信息 | `prs, path` |
| `sanitize_script()` | 检测并修复 Unicode 引号等语法问题 | `path` |

## 目录结构

```
pptx-creator-agent/
├── .github/
│   └── agents/
│       └── pptx-creator.agent.md   # Agent 定义
├── examples/                        # 完整示例
│   ├── README.md                    # 示例说明
│   └── ai-development/              # AI发展演示文稿示例
│       ├── README.md
│       ├── create_presentation.py
│       └── AI-Development-Presentation.pptx
├── packages/                        # 离线安装包
│   ├── README.md                    # 安装说明
│   ├── python_pptx-1.0.2-py3-none-any.whl
│   ├── lxml-6.0.2-cp311-cp311-win_amd64.whl
│   ├── pillow-12.1.1-cp311-cp311-win_amd64.whl
│   ├── typing_extensions-4.15.0-py3-none-any.whl
│   └── xlsxwriter-3.2.9-py3-none-any.whl
├── scripts/
│   ├── pptx_template.py             # 整页级模板库
│   ├── pptx_helpers.py              # 元素级 Helper 函数库
│   ├── example_helpers.py           # pptx_helpers 示例脚本
│   └── example-helpers.pptx         # 示例输出 (8页)
├── requirements.txt                 # 依赖
└── README.md                        # 说明文档
```

## 自定义

### 添加新配色（pptx_template.py）

在 `pptx_template.py` 中添加新的颜色字典：

```python
COLORS_CUSTOM = {
    'bg_dark': RGBColor(0x..., 0x..., 0x...),
    'accent_blue': RGBColor(0x..., 0x..., 0x...),
    'white': RGBColor(0xFF, 0xFF, 0xFF),
    'gray': RGBColor(0xAA, 0xAA, 0xAA),
    'card_bg': RGBColor(0x..., 0x..., 0x...),
    'line_gray': RGBColor(0x..., 0x..., 0x...),
    # 字体配置（必需）
    'font_cn': '微软雅黑',      # 中文正文字体
    'font_en': 'Arial',         # 英文字体
    'font_title_cn': '等线',    # 中文标题字体
}
COLORS = COLORS_CUSTOM
```

### 添加新配色（pptx_helpers.py）

自定义配色只需传入包含 21 个颜色键的字典：

```python
from pptx.dml.color import RGBColor

MY_SCHEME = {
    'navy': RGBColor(0x2C, 0x3E, 0x50),
    'teal': RGBColor(0x16, 0xA0, 0x85),
    'white': RGBColor(0xFF, 0xFF, 0xFF),
    'black': RGBColor(0x2D, 0x34, 0x36),
    # ... 其余键参考 NAVY_TEAL
    'font_cn': '微软雅黑',
    'font_code': 'Courier New',
}
prs, C = create_prs(MY_SCHEME)  # 直接传入 dict
```

### 添加新幻灯片类型

参考现有的 `create_*_slide()` 函数编写新类型。

## 示例

### 整页级模板示例

```bash
cd scripts
python pptx_template.py
```

将生成 `example-presentation.pptx` 示例文件。

### 元素级 Helper 示例

```bash
cd scripts
python example_helpers.py
```

将生成 `example-helpers.pptx`（8 页），展示所有 helper 函数的用法：
- 标题页（`bar`, `tb`, `add_bg`）
- 多行列表（`ml`, `badge`）
- 表格页（`ct`, `ft`, `sc`）
- 代码块（`code`）
- 信息框（`hbox`）
- 富文本列表（`rl`）
- 章节标记（`act_badge`）
- 页码标注（`sn`）

## License

MIT
