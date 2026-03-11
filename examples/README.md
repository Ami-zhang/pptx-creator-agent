# 示例项目

本目录包含使用 PPTX Creator Agent 生成演示文稿的完整示例。

## 示例列表

| 示例 | 说明 | 幻灯片数 |
|------|------|----------|
| [ai-development](ai-development/) | 人工智能发展演示文稿 | 5 页 |

## 如何使用示例

### 1. 查看已生成的 PPTX

每个示例目录都包含已生成的 `.pptx` 文件，可以直接用 PowerPoint 或 WPS 打开查看效果。

### 2. 运行 Python 脚本重新生成

```bash
# 进入示例目录
cd examples/ai-development

# 确保已安装依赖
pip install python-pptx

# 运行脚本
python create_presentation.py
```

### 3. 作为模板修改

复制示例代码到你的项目，根据需要修改：
- 配色方案
- 内容数据
- 幻灯片结构

## 添加新示例

欢迎贡献新的示例！请按以下结构组织：

```
examples/
└── your-example-name/
    ├── README.md              # 示例说明
    ├── create_presentation.py # Python 脚本
    └── output.pptx            # 生成的演示文稿
```
