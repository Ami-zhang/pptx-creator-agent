# AI 发展演示文稿示例

这是一个完整的演示文稿生成示例，展示如何使用 `python-pptx` 创建关于人工智能发展的专业演示文稿。

## 文件说明

| 文件 | 说明 |
|------|------|
| `create_presentation.py` | Python 脚本 - 生成演示文稿的完整代码 |
| `AI-Development-Presentation.pptx` | 输出文件 - 生成的演示文稿 |

## 运行方法

```bash
# 确保已安装 python-pptx
pip install python-pptx

# 运行脚本
python create_presentation.py
```

## 演示文稿内容

生成的演示文稿包含 **5 页幻灯片**：

1. **标题页** - 人工智能发展：从过去到未来
2. **AI发展历程** - 1956-2022+ 关键里程碑时间线
3. **当前AI核心技术** - 6个技术卡片（机器学习、深度学习、NLP、计算机视觉、生成式AI、多模态AI）
4. **AI应用场景** - 6个应用领域（医疗、金融、制造、自动驾驶、教育、内容创作）
5. **未来展望** - 发展趋势与面临挑战

## 设计风格

- **配色方案**: 科技紫蓝风格
  - 背景: 深蓝黑 `#0F0F23`
  - 主强调色: 电子蓝紫 `#667EEA`
  - 辅助色: 紫色 `#764BA2`
  - 点缀色: 青色 `#00D9FF`

- **字体**: Arial (Web-safe)
- **布局**: 16:9 宽屏

## 代码结构

```python
# 配色方案定义
COLORS = {...}

# 工具函数
set_shape_fill()      # 设置形状填充
add_text_frame()      # 添加文本

# 幻灯片创建函数
create_slide1_title()       # 标题页
create_slide2_history()     # 发展历程
create_slide3_technologies() # 核心技术
create_slide4_applications() # 应用场景
create_slide5_future()      # 未来展望

# 主函数
main()  # 组装所有幻灯片并保存
```

## 自定义

你可以基于此示例进行修改：

1. **修改配色**: 编辑 `COLORS` 字典
2. **修改内容**: 编辑各 `create_slide*` 函数中的数据
3. **添加幻灯片**: 参考现有函数创建新的幻灯片类型
