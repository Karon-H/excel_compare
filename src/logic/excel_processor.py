import os
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
            # 自动识别引擎
            ext = os.path.splitext(filepath)[1].lower()
            if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
                engine = 'openpyxl'
            elif ext in ['.xls']:
                engine = 'xlrd'
            else:
                engine = None
            
            # 使用指定引擎读取，确保与后续读取逻辑一致
            xl = pd.ExcelFile(filepath, engine=engine)
            return xl.sheet_names
        except Exception as e:
            # 降级处理：尝试不指定引擎
            try:
                xl = pd.ExcelFile(filepath)
                return xl.sheet_names
            except:
                raise Exception(f"无法获取 Excel Sheet 列表: {e}")

    @staticmethod
    def read_excel_raw(filepath, sheet_name, handle_merged=True):
        """
        极致健壮的 Excel 读取逻辑：
        1. 自动识别引擎，支持多种 Excel 格式
        2. 智能匹配 Sheet 名称（防止空格、大小写导致的失败）
        3. 智能表头识别与合并单元格处理
        """
        # 自动识别引擎
        ext = os.path.splitext(filepath)[1].lower()
        if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            engine = 'openpyxl'
        elif ext in ['.xls']:
            engine = 'xlrd'
        else:
            engine = None
        
        # --- 预处理 Sheet 名称 ---
        # 获取所有实际的 Sheet 名称
        try:
            xl = pd.ExcelFile(filepath, engine=engine)
            actual_sheets = xl.sheet_names
            
            # 1. 尝试完全匹配
            target_sheet = sheet_name
            if sheet_name not in actual_sheets:
                # 2. 尝试模糊匹配 (不区分大小写，不区分首尾空格)
                s_name = str(sheet_name).strip().lower()
                for s in actual_sheets:
                    if s.strip().lower() == s_name:
                        target_sheet = s
                        break
        except Exception as e:
            # 如果获取列表失败，只能盲猜
            target_sheet = sheet_name

        df = None
        
        # --- 尝试读取 ---
        try:
            df = pd.read_excel(filepath, sheet_name=target_sheet, header=None, engine=engine)
        except Exception as e:
            # 如果按名称读取失败，记录错误并抛出，不再默认跳回第一个
            raise Exception(f"读取工作表 '{sheet_name}' 失败: {e}")

        if df is None or df.empty:
            return pd.DataFrame()

        # --- 合并单元格处理 (仅针对 openpyxl) ---
        if handle_merged and engine == 'openpyxl':
            try:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                # 使用匹配后的名称
                ws = None
                if target_sheet in wb.sheetnames:
                    ws = wb[target_sheet]
                else:
                    # 再次兜底匹配
                    match = [s for s in wb.sheetnames if s.lower().strip() == str(target_sheet).lower().strip()]
                    ws = wb[match[0]] if match else wb.worksheets[0]
                
                if ws:
                    for merged_range in ws.merged_cells.ranges:
                        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                        top_left_value = ws.cell(row=min_row, column=min_col).value
                        for r in range(min_row-1, max_row):
                            for c in range(min_col-1, max_col):
                                if r < len(df) and c < len(df.columns):
                                    df.iloc[r, c] = top_left_value
                wb.close()
            except Exception as merge_err:
                print(f"合并单元格处理提示: {merge_err}")
        
        # --- 移除全空行/列 ---
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        if not df.empty:
            # --- 智能表头识别 ---
            header_row_idx = 0
            for i in range(len(df)):
                # 寻找第一个包含非空值的行
                if not df.iloc[i].isna().all():
                    header_row_idx = i
                    break
            
            header_values = df.iloc[header_row_idx]
            df_data = df.iloc[header_row_idx + 1:]
            
            # 处理列名
            new_columns = []
            seen = {}
            for i, val in enumerate(header_values):
                col_name = str(val).strip() if pd.notna(val) and str(val).strip() != "" else f"列{i+1}"
                if col_name in seen:
                    seen[col_name] += 1
                    col_name = f"{col_name}_{seen[col_name]}"
                else:
                    seen[col_name] = 0
                new_columns.append(col_name)
            
            df_data.columns = new_columns
            return df_data.reset_index(drop=True)
        
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
        
        # 校验主键列是否存在
        if key_columns:
            missing_df1 = [col for col in key_columns if col not in df1.columns]
            missing_df2 = [col for col in key_columns if col not in df2.columns]
            if missing_df1:
                raise Exception(f"旧文件中缺少关键列: {', '.join(missing_df1)}")
            if missing_df2:
                raise Exception(f"新文件中缺少关键列: {', '.join(missing_df2)}")

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
        """优化的基于序列对齐的比对逻辑"""
        # 1. 预计算哈希值，提升序列比对和相等判断速度
        # 使用 values 数组操作，避免 iloc/get(col) 产生的 KeyError
        # data_columns 此时包含了新旧文件的列并集，但 df1 或 df2 可能缺失某些列
        
        def get_row_data(df, target_cols):
            # 建立一个与 target_cols 对应的空数据行，确保不会出现 KeyError
            rows = []
            for _, row in df.iterrows():
                row_vals = []
                for col in target_cols:
                    # 使用 row.get 而非 row[col]，因为 df 可能缺少并集中的某些列
                    val = row.get(col, "")
                    row_vals.append(val)
                rows.append(tuple(row_vals))
            return rows

        rows1 = get_row_data(df1, data_columns)
        rows2 = get_row_data(df2, data_columns)
        
        # 2. 序列比对
        matcher = difflib.SequenceMatcher(None, rows1, rows2, autojunk=False)
        results = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # 全等块
                for k in range(i2 - i1):
                    row_vals = list(rows1[i1 + k])
                    results.append((["一致"] + row_vals, ["一致"] + row_vals, ['equal']))
            
            elif tag == 'replace':
                count1 = i2 - i1
                count2 = j2 - j1
                max_count = max(count1, count2)
                for k in range(max_count):
                    has_row1 = k < count1
                    has_row2 = k < count2
                    if has_row1 and has_row2:
                        r1_tuple = rows1[i1 + k]
                        r2_tuple = rows2[j1 + k]
                        
                        if r1_tuple == r2_tuple:
                            vals = list(r1_tuple)
                            results.append((["一致"] + vals, ["一致"] + vals, ['equal']))
                        else:
                            # 逐列比对，标记 >>> diff <<<
                            l_row = []
                            r_row = []
                            has_diff = False
                            for v1, v2 in zip(r1_tuple, r2_tuple):
                                if str(v1).strip() != str(v2).strip():
                                    l_row.append(f">>> {v1} <<<")
                                    r_row.append(f">>> {v2} <<<")
                                    has_diff = True
                                else:
                                    l_row.append(str(v1))
                                    r_row.append(str(v2))
                            
                            status = "修改" if has_diff else "一致"
                            results.append(([status] + l_row, [status] + r_row, ['modified' if has_diff else 'equal']))
                    elif has_row1:
                        row_vals = list(rows1[i1 + k])
                        results.append((["删除"] + row_vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
                    elif has_row2:
                        row_vals = list(rows2[j1 + k])
                        results.append((["新增"] + ["" for _ in data_columns], ["新增"] + row_vals, ['added']))
            
            elif tag == 'delete':
                for k in range(i1, i2):
                    row_vals = list(rows1[k])
                    results.append((["删除"] + row_vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
            
            elif tag == 'insert':
                for k in range(j1, j2):
                    row_vals = list(rows2[k])
                    results.append((["新增"] + ["" for _ in data_columns], ["新增"] + row_vals, ['added']))
                    
        return all_columns, results

    @staticmethod
    def _compare_by_keys(df1, df2, key_columns, data_columns, all_columns):
        """优化的基于主键的比对逻辑"""
        # 建立索引，提高查找速度
        # 使用 set_index 之前，先确保主键列在 df 中存在 (已经在上层 compare_dataframes 校验过)
        # 这里使用 copy 避免修改原始 DataFrame
        d1 = df1.copy()
        d2 = df2.copy()
        
        # 将主键列转换为字符串，避免类型不匹配
        for col in key_columns:
            d1[col] = d1[col].astype(str).str.strip()
            d2[col] = d2[col].astype(str).str.strip()
            
        d1.set_index(key_columns, inplace=True, drop=False)
        d2.set_index(key_columns, inplace=True, drop=False)
        
        # 获取所有主键的并集
        all_keys = list(d1.index.union(d2.index))
        results = []
        
        for key in all_keys:
            in_d1 = key in d1.index
            in_d2 = key in d2.index
            
            if in_d1 and in_d2:
                # 注意：处理重复主键，取第一个
                r1 = d1.loc[[key]].iloc[0]
                r2 = d2.loc[[key]].iloc[0]
                
                # 逐列比对
                l_row = []
                r_row = []
                has_diff = False
                
                for col in data_columns:
                    v1 = r1.get(col, "")
                    v2 = r2.get(col, "")
                    
                    if str(v1).strip() != str(v2).strip():
                        l_row.append(f">>> {v1} <<<")
                        r_row.append(f">>> {v2} <<<")
                        has_diff = True
                    else:
                        l_row.append(str(v1))
                        r_row.append(str(v2))
                
                status = "修改" if has_diff else "一致"
                results.append(([status] + l_row, [status] + r_row, ['modified' if has_diff else 'equal']))
                
            elif in_d1:
                r1 = d1.loc[[key]].iloc[0]
                vals = [str(r1.get(col, "")).strip() for col in data_columns]
                results.append((["删除"] + vals, ["删除"] + ["" for _ in data_columns], ['deleted']))
                
            elif in_d2:
                r2 = d2.loc[[key]].iloc[0]
                vals = [str(r2.get(col, "")).strip() for col in data_columns]
                results.append((["新增"] + ["" for _ in data_columns], ["新增"] + vals, ['added']))
                
        return all_columns, results

    @staticmethod
    def _compare_rows(row1, row2, data_columns):
        """对比单行数据并标记差异"""
        l_row = []
        r_row = []
        has_diff = False
        
        for col in data_columns:
            # 兼容处理：如果 row 中缺少某列，则视为空字符串
            v1 = row1.get(col, "")
            v2 = row2.get(col, "")
            
            # 忽略空白字符的差异
            s1 = str(v1).strip()
            s2 = str(v2).strip()
            
            if s1 != s2:
                l_row.append(f">>> {v1} <<<")
                r_row.append(f">>> {v2} <<<")
                has_diff = True
            else:
                l_row.append(str(v1))
                r_row.append(str(v2))
                
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
