import pandas as pd
import openpyxl
from openpyxl.utils import range_boundaries

class ExcelDiffer:
    """Excel 比对逻辑处理类"""
    
    @staticmethod
    def load_sheets(filepath):
        """读取 Excel 文件的 Sheet 列表"""
        try:
            xl = pd.ExcelFile(filepath)
            return xl.sheet_names
        except Exception as e:
            raise Exception(f"无法读取 Excel 文件: {e}")

    @staticmethod
    def read_excel_raw(filepath, sheet_name, handle_merged=True):
        """
        使用 openpyxl 读取 Excel，保持原始格式和合并单元格信息
        """
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
        ws = wb[sheet_name]
        
        data = []
        # 获取合并单元格范围
        merged_cells = ws.merged_cells.ranges
        
        for row in ws.iter_rows(values_only=True):
            data.append(list(row))
            
        df = pd.DataFrame(data)
        
        if handle_merged:
            # 如果需要处理合并单元格，手动填充合并区域的值
            # openpyxl 的 values_only 模式下，合并单元格只有左上角有值，其余为 None
            # 我们根据 merged_cells 信息进行填充
            for merged_range in merged_cells:
                min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                # 获取左上角的值
                top_left_value = ws.cell(row=min_row, column=min_col).value
                
                # 在 DataFrame 中填充该范围（注意 DataFrame 索引从 0 开始）
                for r in range(min_row-1, max_row):
                    for c in range(min_col-1, max_col):
                        if r < len(df) and c < len(df.columns):
                            df.iloc[r, c] = top_left_value
                            
        # 移除全空的行和列（可选，但通常 Excel 末尾会有很多空行）
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # 将第一行作为列名（如果存在）
        if not df.empty:
            df.columns = [f"列{i+1}" for i in range(len(df.columns))]
            
        return df

    @staticmethod
    def compare_dataframes(df1, df2):
        """
        比对两个 DataFrame，返回比对结果列表。
        返回格式: (all_columns, results)
        results 列表中的每个元素是一个包含 (left_values, right_values, tags) 的元组。
        """
        # 填充缺失值为特定字符串，防止 NaN 比较问题
        df1 = df1.fillna("")
        df2 = df2.fillna("")
        
        results = []
        
        # 获取所有列的并集，并保持原始顺序（优先 df1）
        all_columns = ["状态"] + list(df1.columns)
        for col in df2.columns:
            if col not in all_columns:
                all_columns.append(col)
        
        # 数据列（不含状态列）
        data_columns = all_columns[1:]
        
        # 获取最大行数
        max_rows = max(len(df1), len(df2))
        
        for i in range(max_rows):
            left_values = []
            right_values = []
            tags = []
            
            # 检查行是否存在
            in_df1 = i < len(df1)
            in_df2 = i < len(df2)

            if not in_df1:
                # df1 中没有，df2 中有 (新增行)
                row_data = df2.iloc[i]
                status = "新增"
                left_values = [status] + ["" for _ in data_columns]
                right_values = [status] + [f"{row_data.get(col, '')}" if col in df2.columns else "" for col in data_columns]
                tags.append('added')
            elif not in_df2:
                # df1 中有，df2 中没有 (删除行)
                row_data = df1.iloc[i]
                status = "删除"
                left_values = [status] + [f"{row_data.get(col, '')}" if col in df1.columns else "" for col in data_columns]
                right_values = [status] + ["" for _ in data_columns]
                tags.append('deleted')
            else:
                # 两边都有，逐个单元格比较
                row1 = df1.iloc[i]
                row2 = df2.iloc[i]
                has_diff = False
                l_row = []
                r_row = []
                
                for col in data_columns:
                    val1 = ""
                    val2 = ""
                    
                    if col in df1.columns:
                        val1 = row1[col]
                    if col in df2.columns:
                        val2 = row2[col]
                    
                    # 尝试转换类型比较，避免 int vs float vs str 的问题
                    v1_str = str(val1).strip()
                    v2_str = str(val2).strip()

                    if v1_str != v2_str:
                        l_row.append(f"【{val1}】")
                        r_row.append(f"【{val2}】")
                        has_diff = True
                    else:
                        l_row.append(f"{val1}")
                        r_row.append(f"{val2}")
                
                if has_diff:
                    status = "修改"
                    tags.append('modified')
                else:
                    status = "一致"
                    tags.append('equal')
                
                left_values = [status] + l_row
                right_values = [status] + r_row
            
            results.append((left_values, right_values, tags))
            
        return all_columns, results
