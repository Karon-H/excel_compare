import os
import subprocess
import sys
import logging
import shutil
from .version_manager import VersionManager
from .spec_patcher import SpecPatcher

logger = logging.getLogger("BuildSystem.Engine")

class BuildEngine:
    """核心构建引擎"""

    def __init__(self, config):
        self.config = config
        self.version = "0.0.0"

    def run(self):
        """执行完整构建流程"""
        logger.info("="*30)
        logger.info(f"开始项目构建: {self.config.PROJECT_NAME}")
        logger.info("="*30)

        # 0. 环境检查
        if not self._check_environment():
            return False

        # 1. 提取版本
        self.version = VersionManager.extract(
            self.config.VERSION_FILE, 
            getattr(self.config, 'VERSION_PATTERNS', None)
        )

        # 2. 准备路径
        spec_path = os.path.abspath(self.config.SPEC_FILE)
        dist_dir = os.path.abspath(self.config.DIST_DIR)
        build_dir = os.path.join(os.path.dirname(dist_dir), "build")
        temp_name = f"{self.config.INTERNAL_NAME}_temp"

        # 2.1 清理旧目录
        self._cleanup(dist_dir, build_dir)

        # 3. Spec 打补丁
        if not SpecPatcher.patch_name(spec_path, temp_name):
            return False

        success = False
        try:
            # 4. 运行 PyInstaller
            if self._execute_pyinstaller(spec_path):
                # 5. 重命名产物
                final_exe_name = self.config.OUTPUT_FORMAT.format(
                    project_name=self.config.PROJECT_DISPLAY_NAME,
                    version=self.version
                )
                if self._rename_output(dist_dir, temp_name, final_exe_name):
                    success = True
        finally:
            # 6. 无论成功与否，还原 Spec
            SpecPatcher.revert_name(spec_path, temp_name, self.config.INTERNAL_NAME)

        if success:
            logger.info("="*30)
            logger.info("构建成功完成！")
            logger.info("="*30)
        else:
            logger.error("构建流程在某个环节失败。")
        
        return success

    def _cleanup(self, dist_dir, build_dir):
        """构建前清理"""
        for d in [dist_dir, build_dir]:
            if os.path.exists(d):
                logger.info(f"正在清理目录: {d}")
                try:
                    shutil.rmtree(d)
                except Exception as e:
                    logger.warning(f"无法完全清理目录 {d}: {e}")

    def _check_environment(self):
        """检查构建环境"""
        try:
            subprocess.run(["pyinstaller", "--version"], check=True, capture_output=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            logger.error("未找到 PyInstaller，请先安装: pip install pyinstaller")
            return False

    def _execute_pyinstaller(self, spec_path):
        """调用 PyInstaller 命令行"""
        logger.info(f"正在启动 PyInstaller 打包: {spec_path}")
        try:
            # 使用 subprocess.run 确保实时输出日志
            result = subprocess.run(
                ["pyinstaller", "--noconfirm", spec_path],
                check=True,
                stdout=sys.stdout,
                stderr=sys.stderr,
                text=True
            )
            return result.returncode == 0
        except subprocess.CalledProcessError as e:
            logger.error(f"PyInstaller 执行失败: {e}")
            return False
        except Exception as e:
            logger.error(f"启动 PyInstaller 时发生异常: {e}")
            return False

    def _rename_output(self, dist_dir, temp_name, final_name):
        """重命名生成的可执行文件"""
        temp_exe = os.path.join(dist_dir, f"{temp_name}.exe")
        final_exe = os.path.join(dist_dir, f"{final_name}.exe")

        try:
            if not os.path.exists(temp_exe):
                logger.error(f"找不到生成的临时文件: {temp_exe}")
                return False

            if os.path.exists(final_exe):
                os.remove(final_exe)
                logger.info(f"清理旧的产物: {final_exe}")

            shutil.move(temp_exe, final_exe)
            logger.info(f"产物已就绪: {final_exe}")
            return True
        except Exception as e:
            logger.error(f"重命名产物失败: {e}")
            return False
