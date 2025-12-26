import os

class BuildSettings:
    """项目构建配置 - 移植到新工程时仅需修改此处"""

    # 1. 项目基础信息
    PROJECT_NAME = "ExcelCompareTool"
    PROJECT_DISPLAY_NAME = "Excel差异比对工具"
    INTERNAL_NAME = "ExcelCompareTool"  # Spec 文件中 EXE 对象的 name 字段

    # 2. 路径配置
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    VERSION_FILE = os.path.join(ROOT_DIR, "..", "updates_notes.txt")
    SPEC_FILE = os.path.join(ROOT_DIR, "..", "build_system", "templates", "excel_compare.spec")
    DIST_DIR = os.path.join(ROOT_DIR, "..", "dist")

    # 3. 输出格式
    # 可用的变量: {project_name}, {version}
    OUTPUT_FORMAT = "{project_name}{version}"

    # 4. (可选) 自定义版本匹配正则
    # 匹配 updates_notes.txt 中的 V2.44 这种格式
    VERSION_PATTERNS = [
        r"(V\d+\.\d+(?:\.\d+)?)",
        r"VERSION\s*=\s*['\"]([^'\"]+)['\"]"
    ]
