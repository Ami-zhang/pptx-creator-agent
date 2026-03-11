# 离线安装包

本目录包含 `python-pptx` 及其依赖的离线安装包（.whl 文件），方便在无网络环境下安装。

## 包列表

| 包名 | 版本 | 文件大小 | 适用平台 |
|------|------|----------|----------|
| python-pptx | 1.0.2 | 462 KB | 任意平台 (py3-none-any) |
| lxml | 6.0.2 | 3.9 MB | Windows 64位, Python 3.11 |
| Pillow | 12.1.1 | 6.8 MB | Windows 64位, Python 3.11 |
| typing_extensions | 4.15.0 | 44 KB | 任意平台 (py3-none-any) |
| XlsxWriter | 3.2.9 | 172 KB | 任意平台 (py3-none-any) |

## 系统环境要求

### 当前包适用环境

| 项目 | 要求 |
|------|------|
| **操作系统** | Windows 10/11 64位 |
| **Python 版本** | Python 3.11.x |
| **架构** | AMD64 (x86_64) |

> ⚠️ **注意**: `lxml` 和 `Pillow` 是平台相关的包，当前提供的版本仅适用于 **Windows 64位 + Python 3.11**。其他系统需要从 PyPI 在线安装或下载对应版本。

## 安装方法

### 方法 1: 一键安装所有包（推荐）

```bash
# 进入 packages 目录
cd packages

# 安装所有 whl 文件
pip install *.whl
```

### 方法 2: 按顺序安装

```bash
cd packages

# 1. 先安装依赖包
pip install typing_extensions-4.15.0-py3-none-any.whl
pip install lxml-6.0.2-cp311-cp311-win_amd64.whl
pip install pillow-12.1.1-cp311-cp311-win_amd64.whl
pip install xlsxwriter-3.2.9-py3-none-any.whl

# 2. 最后安装 python-pptx
pip install python_pptx-1.0.2-py3-none-any.whl
```

### 方法 3: 在线安装（其他系统）

如果你使用的是 **macOS**、**Linux** 或 **其他 Python 版本**，请使用在线安装：

```bash
pip install python-pptx
```

或者手动下载适合你系统的包：
- PyPI: https://pypi.org/project/python-pptx/
- lxml: https://pypi.org/project/lxml/
- Pillow: https://pypi.org/project/Pillow/

## 验证安装

安装完成后，运行以下命令验证：

```python
python -c "from pptx import Presentation; print('python-pptx 安装成功!')"
```

## 其他系统的包下载

如果需要其他系统的离线包，可以从以下地址下载：

| 系统 | lxml 下载 | Pillow 下载 |
|------|-----------|-------------|
| macOS (Intel) | [lxml](https://pypi.org/project/lxml/#files) | [Pillow](https://pypi.org/project/Pillow/#files) |
| macOS (Apple Silicon) | 同上 | 同上 |
| Linux x86_64 | 同上 | 同上 |
| Windows 32位 | 同上 | 同上 |

选择对应你 Python 版本的 `.whl` 文件下载即可。

## 包文件命名规则

`.whl` 文件名格式：`{包名}-{版本}-{python版本}-{abi}-{平台}.whl`

示例解读：
- `lxml-6.0.2-cp311-cp311-win_amd64.whl`
  - `cp311` = CPython 3.11
  - `win_amd64` = Windows 64位
  
- `python_pptx-1.0.2-py3-none-any.whl`
  - `py3` = Python 3.x 通用
  - `none` = 无 ABI 要求
  - `any` = 任意平台
