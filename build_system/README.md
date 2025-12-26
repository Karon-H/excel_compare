# Standardized Portable Build Framework
# 标准化可移植打包框架

这是一个基于 Python 驱动的 PyInstaller 增强打包框架，旨在解决 Windows 环境下打包时的路径、编码、版本管理及产物重命名等问题。

## 1. 目录结构
- `builder.py`: 构建入口脚本。
- `config.py`: 项目配置文件（移植时仅需修改此文件）。
- `core/`: 框架核心逻辑（版本管理、Spec 修正、构建引擎）。
- `templates/`: 包含可复用的 `.spec` 模板。

## 2. 核心功能
1.  **版本自动同步**: 自动从 Markdown 或 Python 文件中提取语义化版本号。
2.  **Spec 动态补丁**: 构建期间自动修改 Spec 文件，规避中文路径导致的 PyInstaller 崩溃问题。
3.  **产物标准化**: 自动清理旧产物，并按 `{项目名}{版本号}.exe` 格式重命名新产物。
4.  **环境自动还原**: 无论构建成功或失败，都会将配置文件还原，保持 Git 工作区整洁。

## 3. 移植指南 (How to Port)
将 `build_system/` 目录拷贝到您的新项目中，然后按照以下步骤配置：

### 第一步：修改 `config.py`
根据新项目的信息修改以下参数：
```python
PROJECT_NAME = "MyNewProject"
PROJECT_DISPLAY_NAME = "我的新程序"
INTERNAL_NAME = "main"  # 对应 Spec 文件中 EXE 对象的 name 字段
VERSION_FILE = "../version.txt" # 存放版本号的文件路径
SPEC_FILE = "../main.spec"     # Spec 文件路径
```

### 第二步：准备版本文件
在 `VERSION_FILE` 指定的文件中包含类似 `版本 1.0.0` 或 `v1.0.0` 的文字。

### 第三步：运行构建
在终端执行：
```bash
python build_system/builder.py
```

## 4. 优势
- **解耦**: 构建逻辑与业务代码完全分离。
- **配置化**: 无需修改任何脚本逻辑即可适配新项目。
- **强健性**: 自动化处理了大部分打包过程中的手工易错环节。
