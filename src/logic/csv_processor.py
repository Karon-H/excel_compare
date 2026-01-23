import os
import pandas as pd
from src.logic.base_processor import BaseDiffer
from src.logic.excel_processor import ExcelDiffer

class CSVDiffer(BaseDiffer):
    """CSV 比对逻辑处理类 (插件示例)"""

    def get_containers(self, filepath):
        """CSV 文件只有一个容器，即文件本身"""
        return ["Default"]

    def read_data(self, filepath, container_name, **kwargs):
        """读取 CSV 数据"""
        try:
            # 尝试多种编码
            for encoding in ['utf-8', 'gbk', 'utf-8-sig']:
                try:
                    df = pd.read_csv(filepath, encoding=encoding)
                    return df
                except:
                    continue
            raise Exception("无法识别 CSV 编码")
        except Exception as e:
            raise Exception(f"读取 CSV 失败: {e}")

    def compare(self, df1, df2, key_columns=None, **kwargs):
        """复用 ExcelDiffer 的比对逻辑 (因为核心是 DataFrame)"""
        return ExcelDiffer.compare_dataframes(df1, df2, key_columns)

    def export(self, filepath, columns, results, key_columns=None, **kwargs):
        """CSV 暂不支持带格式导出，降级为普通 CSV 导出或复用 Excel 导出"""
        # 为了演示，我们直接使用 ExcelDiffer 的导出，虽然是 CSV 插件，但可以导出为 XLSX
        if filepath.lower().endswith('.csv'):
            # 简化的 CSV 导出逻辑
            data = []
            for left, right, info in results:
                status = info.get('status', '')
                row = [status] + list(right if status != 'deleted' else left)
                data.append(row)
            df = pd.DataFrame(data, columns=['Status'] + list(columns))
            df.to_csv(filepath, index=False, encoding='utf-8-sig')
        else:
            return ExcelDiffer.export_diff(filepath, columns, results, key_columns)

    def compare_all(self, file1, file2, key_columns=None, progress_callback=None, **kwargs):
        """CSV 文件的批量比对即单表比对"""
        if progress_callback:
            progress_callback(50, "正在比对 CSV 文件...")
        
        try:
            df1 = self.read_data(file1, "Default")
            df2 = self.read_data(file2, "Default")
            cols, results = self.compare(df1, df2, key_columns)
            
            stats = {'added': 0, 'deleted': 0, 'modified': 0, 'equal': 0}
            for r in results:
                status = r[2].get('status', 'equal')
                if status in stats:
                    stats[status] += 1
            
            return [{
                'sheet1': os.path.basename(file1),
                'sheet2': os.path.basename(file2),
                'stats': stats,
                'status': 'success'
            }]
        except Exception as e:
            return [{
                'sheet1': os.path.basename(file1),
                'sheet2': os.path.basename(file2),
                'error': str(e),
                'status': 'error'
            }]
