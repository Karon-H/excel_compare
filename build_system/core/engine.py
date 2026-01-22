import os
import subprocess
import sys
import logging
import shutil
import string
import ctypes
from .version_manager import VersionManager
from .spec_patcher import SpecPatcher

logger = logging.getLogger("BuildSystem.Engine")

def get_available_drive():
    """获取 Windows 系统中未使用的盘符"""
    if sys.platform != 'win32':
        return None
    
    # 尝试从 Z 倒着找
    import ctypes
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for letter in reversed(string.ascii_uppercase):
        if not (bitmask & (1 << (ord(letter) - ord('A')))):
            return f"{letter}:"
    return None

class BuildEngine:
    """核心构建引擎"""

    def __init__(self, config):
        self.config = config
        self.version = "0.0.0"
        self.virtual_drive = None
        self.original_cwd = os.getcwd()

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

        # 2. 检查路径是否包含中文，如果包含则使用 subst 映射
        project_root = os.path.abspath(os.path.join(self.config.ROOT_DIR, ".."))
        needs_subst = False
        try:
            project_root.encode('ascii')
        except UnicodeEncodeError:
            needs_subst = True

        if needs_subst and sys.platform == 'win32':
            self.virtual_drive = get_available_drive()
            if self.virtual_drive:
                logger.info(f"检测到中文路径，正在将项目映射到虚拟盘符: {self.virtual_drive}")
                try:
                    subprocess.run(["subst", self.virtual_drive, project_root], check=True)
                    # 切换到虚拟盘符下的对应路径
                    os.chdir(self.virtual_drive)
                    logger.info(f"成功切换到虚拟盘符: {os.getcwd()}")
                except Exception as e:
                    logger.error(f"映射虚拟盘符失败: {e}")
                    return False
            else:
                logger.warning("未找到可用的虚拟盘符，尝试直接构建...")

        def to_virtual(path):
            if self.virtual_drive and path.lower().startswith(project_root.lower()):
                return os.path.join(self.virtual_drive, path[len(project_root):].lstrip("\\/"))
            return path

        success = False
        try:
            # 3. 准备路径 (此时如果是虚拟盘符，路径需要转换为虚拟盘路径)
            dist_dir = to_virtual(os.path.abspath(self.config.DIST_DIR))
            build_dir = os.path.join(os.path.dirname(dist_dir), "build")
            spec_path = to_virtual(os.path.abspath(self.config.SPEC_FILE))
            temp_name = f"{self.config.INTERNAL_NAME}_temp"

            # 3.1 清理旧目录
            self._cleanup(dist_dir, build_dir)

            # 4. Spec 打补丁
            if not SpecPatcher.patch_name(spec_path, temp_name):
                return False

            # 5. 运行 PyInstaller
            if self._execute_pyinstaller(spec_path, to_virtual):
                # 6. 重命名产物
                final_exe_name = self.config.OUTPUT_FORMAT.format(
                    project_name=self.config.PROJECT_DISPLAY_NAME,
                    version=self.version
                )
                if self._rename_output(dist_dir, temp_name, final_exe_name):
                    success = True
            
            # 还原 Spec
            SpecPatcher.revert_name(spec_path, temp_name, self.config.INTERNAL_NAME)
        finally:
            # 7. 清理虚拟盘符
            if self.virtual_drive:
                logger.info(f"正在解除虚拟盘符映射: {self.virtual_drive}")
                os.chdir(self.original_cwd)
                subprocess.run(["subst", self.virtual_drive, "/d"], capture_output=True)

        if success:
            logger.info("="*30)
            logger.info("构建成功完成！")
            logger.info("="*30)
        else:
            logger.error("构建流程在某个环节失败。")
        
        return success

    def _cleanup(self, dist_dir, build_dir):
        """构建前清理"""
        # 1. 强制清理 build 目录，这是 PyInstaller 的临时中间件
        if os.path.exists(build_dir):
            logger.info(f"正在清理构建缓存目录: {build_dir}")
            try:
                shutil.rmtree(build_dir)
            except Exception as e:
                logger.warning(f"无法完全清理目录 {build_dir}: {e}")
        
        # 2. 不再删除整个 dist 目录，以保留历史版本
        if not os.path.exists(dist_dir):
            os.makedirs(dist_dir)
        else:
            logger.info(f"保持输出目录: {dist_dir} (保留历史版本)")

    def _check_environment(self):
        """检查构建环境"""
        try:
            subprocess.run(["pyinstaller", "--version"], check=True, capture_output=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            logger.error("未找到 PyInstaller，请先安装: pip install pyinstaller")
            return False

    def _execute_pyinstaller(self, spec_path, to_virtual_fn):
        """调用 PyInstaller 命令行"""
        logger.info(f"正在启动 PyInstaller 打包: {spec_path}")
        try:
            # 使用 python -m PyInstaller 确保使用当前环境，并转换解释器路径为虚拟路径
            python_exe = to_virtual_fn(sys.executable)
            cmd = [python_exe, "-m", "PyInstaller", "--noconfirm", spec_path]
            
            # 使用 subprocess.run 确保实时输出日志
            result = subprocess.run(
                cmd,
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
                # 检查是否生成了目录而不是单文件 (PyInstaller 默认行为可能受环境影响)
                temp_dir = os.path.join(dist_dir, temp_name)
                if os.path.exists(temp_dir):
                    # 如果是目录，则重命名目录
                    final_dir = os.path.join(dist_dir, final_name)
                    if os.path.exists(final_dir):
                        shutil.rmtree(final_dir)
                    shutil.move(temp_dir, final_dir)
                    logger.info(f"产物目录已就绪: {final_dir}")
                    return True
                
                logger.error(f"找不到生成的临时文件或目录: {temp_exe}")
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
