import pandas as pd
import openpyxl
import difflib
from openpyxl.utils import range_boundaries
from openpyxl.styles import PatternFill, Font, Alignment

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
    def compare_dataframes(df1, df2, key_columns=None):
        """
        比对两个 DataFrame。
        - 如果 key_columns 为空，使用序列对齐算法 (difflib.SequenceMatcher)。
        - 如果 key_columns 不为空，使用主键比对算法。
        返回格式: (all_columns, results)
        """
        # 填充缺失值为特定字符串
        df1 = df1.fillna("")
        df2 = df2.fillna("")
        
        # 获取所有列的并集
        all_columns = ["状态"] + list(df1.columns)
        for col in df2.columns:
            if col not in all_columns:
                all_columns.append(col)
        data_columns = all_columns[1:]

        if not key_columns:
            return ExcelDiffer._compare_by_sequence(df1, df2, data_columns, all_columns)
        else:
            return ExcelDiffer._compare_by_keys(df1, df2, key_columns, data_columns, all_columns)

    @staticmethod
    def _compare_by_sequence(df1, df2, data_columns, all_columns):
        """原有的基于序列对齐的比对逻辑"""
        # 将 DataFrame 转换为字符串列表进行序列比对
        rows1 = [tuple(row) for row in df1.values]
        rows2 = [tuple(row) for row in df2.values]
        
        matcher = difflib.SequenceMatcher(None, rows1, rows2)
        results = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for k in range(i2 - i1):
                    row = df1.iloc[i1 + k]
                    vals = [row.get(col, "") for col in data_columns]
                    results.append((["一致"] + vals, ["一致"] + vals, ['equal']))
            
            elif tag == 'replace':
                count1 = i2 - i1
                count2 = j2 - j1
                max_count = max(count1, count2)
                for k in range(max_count):
                    has_row1 = k < count1
                    has_row2 = k < count2
                    if has_row1 and has_row2:
                        row1 = df1.iloc[i1 + k]
                        row2 = df2.iloc[j1 + k]
                        l_row, r_row, has_diff = ExcelDiffer._compare_rows(row1, row2, data_columns)
                        status = "修改" if has_diff else "一致"
                        results.append(([status] + l_row, [status] + r_row, ['modified' if has_diff else 'equal']))
                    elif has_row1:
                        row1 = df1.iloc[i1 + k]
                        vals = [row1.get(col, "") for col in data_columns]
                        results.append((["删除"] + vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
                    elif has_row2:
                        row2 = df2.iloc[j1 + k]
                        vals = [row2.get(col, "") for col in data_columns]
                        results.append((["新增"] + ["" for _ in data_columns], ["新增"] + vals, ['added']))
            
            elif tag == 'delete':
                for k in range(i1, i2):
                    row = df1.iloc[k]
                    vals = [row.get(col, "") for col in data_columns]
                    results.append((["删除"] + vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
            
            elif tag == 'insert':
                for k in range(j1, j2):
                    row = df2.iloc[k]
                    vals = [row.get(col, "") for col in data_columns]
                    results.append((["新增"] + ["" for _ in data_columns], ["新增"] + vals, ['added']))
                    
        return all_columns, results

    @staticmethod
    def _compare_by_keys(df1, df2, key_columns, data_columns, all_columns):
        """基于关键列的主键比对逻辑"""
        # 1. 为两个 DF 创建索引字典
        def get_key(row, keys):
            return tuple(str(row.get(k, "")) for k in keys)

        dict1 = {get_key(row, key_columns): row for _, row in df1.iterrows()}
        dict2 = {get_key(row, key_columns): row for _, row in df2.iterrows()}
        
        # 2. 获取所有的 Key 集合，并保持一定的顺序（以文件2为主，结合文件1）
        all_keys = []
        keys1 = list(dict1.keys())
        keys2 = list(dict2.keys())
        
        # 简单的合并策略：先按文件1的顺序放，再放文件2中新增的
        all_keys = keys1.copy()
        set_keys1 = set(keys1)
        for k in keys2:
            if k not in set_keys1:
                all_keys.append(k)
        
        results = []
        for k in all_keys:
            in_1 = k in dict1
            in_2 = k in dict2
            
            if in_1 and in_2:
                # 匹配到主键，比对内容
                row1 = dict1[k]
                row2 = dict2[k]
                l_row, r_row, has_diff = ExcelDiffer._compare_rows(row1, row2, data_columns)
                status = "修改" if has_diff else "一致"
                results.append(([status] + l_row, [status] + r_row, ['modified' if has_diff else 'equal']))
            elif in_1:
                # 只有文件1有 -> 删除
                row1 = dict1[k]
                vals = [row1.get(col, "") for col in data_columns]
                results.append((["删除"] + vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
            elif in_2:
                # 只有文件2有 -> 新增
                row2 = dict2[k]
                vals = [row2.get(col, "") for col in data_columns]
                results.append((["新增"] + ["" for _ in data_columns], ["新增"] + vals, ['added']))
                
        return all_columns, results

    @staticmethod
    def _compare_rows(row1, row2, columns):
        """内部方法：比对两行数据并标记差异"""
        l_row = []
        r_row = []
        has_diff = False
        
        for col in columns:
            val1 = row1.get(col, "")
            val2 = row2.get(col, "")
            
            v1_str = str(val1).strip()
            v2_str = str(val2).strip()
            
            if v1_str != v2_str:
                l_row.append(f">>> {val1} <<<")
                r_row.append(f">>> {val2} <<<")
                has_diff = True
            else:
                l_row.append(str(val1))
                r_row.append(str(val2))
        return l_row, r_row, has_diff


    @staticmethod
    def export_diff(output_path, columns, results):
        """将比对结果导出到 Excel，并应用单元格级别的高亮"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "比对结果"

        # 定义样式
        header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        modified_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        added_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        deleted_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal='center', vertical='center')

        # 写入表头 (左右两部分)
        half_cols = len(columns)
        for i, col in enumerate(columns):
            # 左侧表头
            cell_l = ws.cell(row=1, column=i+1, value=f"文件1_{col}")
            cell_l.fill = header_fill
            cell_l.font = bold_font
            # 右侧表头
            cell_r = ws.cell(row=1, column=i+half_cols+2, value=f"文件2_{col}")
            cell_r.fill = header_fill
            cell_r.font = bold_font

        # 写入数据并设置高亮
        for row_idx, (left_vals, right_vals, tags) in enumerate(results):
            excel_row = row_idx + 2
            
            # 处理左侧数据
            for col_idx, val in enumerate(left_vals):
                val_str = str(val)
                actual_val = val
                is_diff = False
                
                if val_str.startswith(">>> ") and val_str.endswith(" <<<"):
                    actual_val = val_str[4:-4]
                    is_diff = True
                
                cell = ws.cell(row=excel_row, column=col_idx+1, value=actual_val)
                
                if 'deleted' in tags:
                    cell.fill = deleted_fill
                elif is_diff:
                    cell.fill = modified_fill

            # 处理右侧数据
            for col_idx, val in enumerate(right_vals):
                val_str = str(val)
                actual_val = val
                is_diff = False
                
                if val_str.startswith(">>> ") and val_str.endswith(" <<<"):
                    actual_val = val_str[4:-4]
                    is_diff = True
                
                cell = ws.cell(row=excel_row, column=col_idx+half_cols+2, value=actual_val)
                
                if 'added' in tags:
                    cell.fill = added_fill
                elif is_diff:
                    cell.fill = modified_fill

        # 自动调整列宽
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = min(adjusted_width, 50)

        wb.save(output_path)
