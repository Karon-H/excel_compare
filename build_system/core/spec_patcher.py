import logging
import os
import re

logger = logging.getLogger("BuildSystem.Spec")

class SpecPatcher:
    """Spec 文件修正工具，用于在构建期间动态修改配置"""

    @staticmethod
    def patch_name(spec_path, temp_name):
        """修改 Spec 文件中的 name 字段 (仅限 EXE 对象)"""
        try:
            if not os.path.exists(spec_path):
                logger.error(f"Spec 文件不存在: {spec_path}")
                return False

            with open(spec_path, "r", encoding="utf-8") as f:
                content = f.read()

            # 使用正则精准匹配 EXE 括号内的 name='...'
            # 匹配模式：EXE( 后面跟任意字符，直到碰到 name='...' 或 name="..."
            pattern = r"(EXE\([\s\S]*?name\s*=\s*)['\"].*?['\"]"
            replacement = r"\1" + f"'{temp_name}'"
            
            new_content, count = re.subn(pattern, replacement, content)

            if count == 0:
                logger.warning("未在 Spec 文件中的 EXE 对象内找到 name 字段")
                return False

            with open(spec_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            
            logger.info(f"Spec 文件已打补丁: EXE name -> {temp_name}")
            return True
        except Exception as e:
            logger.error(f"打补丁失败: {e}")
            return False

    @staticmethod
    def revert_name(spec_path, temp_name, original_name):
        """还原 Spec 文件中的 name 字段"""
        try:
            with open(spec_path, "r", encoding="utf-8") as f:
                content = f.read()
            
            old_str = f"name='{temp_name}',"
            new_str = f"name='{original_name}',"
            
            if old_str in content:
                content = content.replace(old_str, new_str)
                with open(spec_path, "w", encoding="utf-8") as f:
                    f.write(content)
                logger.info(f"Spec 文件已还原: {original_name}")
                return True
            else:
                logger.warning(f"在 Spec 中未找到占位符 {old_str}，无法还原")
                return False
        except Exception as e:
            logger.error(f"还原 Spec 失败: {e}")
            return False
