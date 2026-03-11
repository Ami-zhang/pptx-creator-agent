# PPTX Creator Agent

一个 VS Code Copilot 自定义 Agent，用于从文本内容自动生成专业的 PowerPoint 演示文稿。

## 功能特点

- 🎨 **智能设计** - 根据内容主题自动选择配色方案
- 📝 **多种输入** - 支持直接文本、Markdown 文件等
- 🎯 **专业模板** - 提供多种幻灯片布局（标题页、内容页、卡片页、总结页）
- 🔧 **易于扩展** - 模块化的 Python 代码，方便自定义

## 安装

### 1. 安装依赖

```bash
pip install python-pptx
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
    create_summary_slide
)
from pptx import Presentation
from pptx.util import Inches

# 创建演示文稿
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# 添加幻灯片
create_title_slide(prs, "标题", "副标题")
create_content_slide(prs, "内容", [("要点", "说明")])

# 保存
prs.save("output.pptx")
```

## 配色方案

内置两种配色方案：

### 科技主题 (COLORS_TECH)
- 深蓝黑背景
- 电子蓝紫强调色
- 青色点缀

### 商务主题 (COLORS_BUSINESS)
- 深灰蓝背景
- 金色强调
- 专业稳重

可以在 `pptx_template.py` 中修改 `COLORS = COLORS_TECH` 来切换。

## 幻灯片类型

| 类型 | 函数 | 用途 |
|------|------|------|
| 标题页 | `create_title_slide()` | 演示文稿首页 |
| 内容页 | `create_content_slide()` | 列表形式的要点 |
| 卡片页 | `create_cards_slide()` | 网格布局，最多6个卡片 |
| 总结页 | `create_summary_slide()` | 结束语、感谢页 |

## 目录结构

```
pptx-creator-agent/
├── .github/
│   └── agents/
│       └── pptx-creator.agent.md   # Agent 定义
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
    # ...
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
