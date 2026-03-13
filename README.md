# PPTX Creator Agent

一个 VS Code Copilot 自定义 Agent，用于从文本内容自动生成专业的 PowerPoint 演示文稿。

## 功能特点

- 🎨 **智能设计** - 根据内容主题自动选择配色方案（科技/商务/企业浅色）
- 📝 **多种输入** - 支持直接文本、Markdown 文件、现有 PPTX 模板等
- 🎯 **专业模板** - 提供多种幻灯片布局（标题页、内容页、卡片页、表格页、总结页）
- 📊 **表格支持** - 创建表格、合并单元格、自定义样式，满足企业 PPT 需求
- 🀄 **中文友好** - 内置微软雅黑/等线等 CJK 字体配置，告别 Arial 乱码
- 🔍 **模板分析** - 分析现有 PPTX 结构（shapes、表格布局、样式），辅助复刻
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

## 配色方案

内置三种配色方案：

### 科技主题 (COLORS_TECH) — 深色
- 深蓝黑背景 (#0F0F23)
- 电子蓝紫强调色
- 青色点缀

### 商务主题 (COLORS_BUSINESS) — 深色
- 深灰蓝背景 (#1C2833)
- 金色强调
- 专业稳重

### 企业主题 (COLORS_CORPORATE) — 浅色 ⭐
- 白色背景 + 蓝色表头 (#2E5C9A)
- 黑色正文，适合大多数企业场景
- 内置等线（标题）+ 微软雅黑（正文）字体

可以在 `pptx_template.py` 中修改 `COLORS = COLORS_CORPORATE` 来切换。

每种配色方案均包含字体配置（`font_cn`/`font_en`/`font_title_cn`），确保中英文混排正确渲染。

## 幻灯片类型

| 类型 | 函数 | 用途 |
|------|------|------|
| 标题页 | `create_title_slide()` | 演示文稿首页 |
| 内容页 | `create_content_slide()` | 列表形式的要点 |
| 卡片页 | `create_cards_slide()` | 网格布局，最多6个卡片 |
| 表格页 | `create_table_slide()` | 带样式表头的数据表格 |
| 总结页 | `create_summary_slide()` | 结束语、感谢页 |

### 表格工具函数

| 函数 | 用途 |
|------|------|
| `create_table()` | 底层表格创建 |
| `set_cell()` | 设置单元格内容/样式（支持多行文本） |
| `merge_cells()` | 合并单元格（需先合并再写内容） |
| `set_table_style()` | 设置表格样式（通过 XML） |

### 布局辅助

| 函数 | 用途 |
|------|------|
| `add_section_label()` | 分节标签（如蓝底白字横条） |
| `analyze_template()` | 分析现有 PPTX 模板结构 |

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
│   └── pptx_template.py            # Python 模板库
├── requirements.txt                 # 依赖
└── README.md                        # 说明文档
```

## 自定义

### 添加新配色

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

### 添加新幻灯片类型

参考现有的 `create_*_slide()` 函数编写新类型。

## 示例

运行示例脚本：

```bash
cd scripts
python pptx_template.py
```

将生成 `example-presentation.pptx` 示例文件。

## License

MIT
