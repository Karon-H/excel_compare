import re
import os
import logging

logger = logging.getLogger("BuildSystem.Version")

class VersionManager:
    """版本管理工具，负责从文件中提取版本号"""
    
    DEFAULT_PATTERNS = [
        r"版本[:：]\s*(\d+\.\d+\.\d+)",
        r"版本\s*(\d+\.\d+\.\d+)",
        r"v(\d+\.\d+\.\d+)",
        r"VERSION\s*=\s*['\"](\d+\.\d+\.\d+)['\"]"
    ]

    @staticmethod
    def extract(file_path, patterns=None):
        """从指定文件提取版本号"""
        if not os.path.exists(file_path):
            logger.warning(f"版本文件未找到: {file_path}")
            return "0.0.0"
        
        search_patterns = patterns or VersionManager.DEFAULT_PATTERNS
        
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
            
            for pattern in search_patterns:
                match = re.search(pattern, content, flags=re.IGNORECASE)
                if match:
                    version = match.group(1)
                    logger.info(f"成功提取版本号: {version}")
                    return version
        except Exception as e:
            logger.error(f"提取版本号时出错: {e}")
            
        return "Unknown"
