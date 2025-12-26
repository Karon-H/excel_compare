#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标准化构建框架入口
"""

import os
import sys
import logging

# 初始化日志
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s][%(levelname)s][%(name)s] %(message)s',
    datefmt='%H:%M:%S'
)

# 确保可以导入核心模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from config import BuildSettings
    from core.engine import BuildEngine
except ImportError as e:
    print(f"导入失败: {e}")
    sys.exit(1)

def main():
    # 确保在项目根目录下运行
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    os.chdir(project_root)
    
    config = BuildSettings()
    engine = BuildEngine(config)
    
    success = engine.run()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
