# 示例项目

本目录包含使用 PPTX Creator Agent 生成演示文稿的完整示例。

## 示例列表

| 示例 | 说明 | 使用库 | 幻灯片数 |
|------|------|--------|----------|
| [ai-development/](ai-development/) | 人工智能发展演示文稿（独立脚本） | python-pptx 直接调用 | 5 页 |
| [example_helpers.py](example_helpers.py) | pptx_helpers 全功能演示 | `scripts/pptx_helpers.py` | 8 页 |
| [validate_15page_deck.py](validate_15page_deck.py) | 16 页分块生成验证 | `scripts/pptx_helpers.py` | 16 页 |
| [example-presentation.pptx](example-presentation.pptx) | pptx_template 整页模板示例输出 | `scripts/pptx_template.py` | — |

## 如何使用示例

### 1. 查看已生成的 PPTX

本目录包含已生成的 `.pptx` 文件，可以直接用 PowerPoint 或 WPS 打开查看效果：

- `example-helpers.pptx` — 8 页 helper 函数演示
- `validate-15page-deck.pptx` — 16 页分块生成验证
- `example-presentation.pptx` — 整页级模板示例
- `ai-development/AI-Development-Presentation.pptx` — 独立 AI 发展演示

### 2. 运行 Python 脚本重新生成

```bash
# 确保已安装依赖
pip install python-pptx

# Helper 全功能演示（8 页）
cd examples
python example_helpers.py

# 分块生成验证（16 页）
python validate_15page_deck.py

# AI 发展演示（独立脚本，5 页）
cd ai-development
python create_presentation.py
```

> **注意**: `example_helpers.py` 和 `validate_15page_deck.py` 依赖 `scripts/pptx_helpers.py`，请从项目根目录运行或确保路径正确。

### 3. 作为模板修改

复制示例代码到你的项目，根据需要修改：
- 配色方案
- 内容数据
- 幻灯片结构
