import unittest
import pandas as pd
import os
import shutil
from src.logic.excel_processor import ExcelDiffer

class TestExcelDiffer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """准备测试用的 Excel 文件"""
        cls.test_dir = "test_tmp"
        if not os.path.exists(cls.test_dir):
            os.makedirs(cls.test_dir)
            
        cls.file1 = os.path.join(cls.test_dir, "test1.xlsx")
        cls.file2 = os.path.join(cls.test_dir, "test2.xlsx")
        cls.file_empty = os.path.join(cls.test_dir, "empty.xlsx")
        
        # 创建标准测试数据
        df1 = pd.DataFrame({
            "ID": [1, 2, 3],
            "Name": ["Alice", "Bob", "Charlie"],
            "Age": [25, 30, 35]
        })
        df1.to_excel(cls.file1, index=False)
        
        # 创建有差异的测试数据 (用于主键比对)
        # 1. 修改 ID=2 的 Age
        # 2. 删除 ID=3
        # 3. 新增 ID=4
        df2 = pd.DataFrame({
            "ID": [1, 2, 4],
            "Name": ["Alice", "Bob", "David"],
            "Age": [25, 31, 40]
        })
        df2.to_excel(cls.file2, index=False)
        
        # 创建空文件 (只有表头)
        df_empty = pd.DataFrame(columns=["A", "B", "C"])
        df_empty.to_excel(cls.file_empty, index=False)

    @classmethod
    def tearDownClass(cls):
        """清理测试文件"""
        if os.path.exists(cls.test_dir):
            shutil.rmtree(cls.test_dir)

    def test_load_sheets(self):
        """测试获取 Sheet 列表"""
        sheets = ExcelDiffer.load_sheets(self.file1)
        self.assertIn("Sheet1", sheets)

    def test_read_excel_raw(self):
        """测试读取 Excel 数据"""
        df = ExcelDiffer.read_excel_raw(self.file1, "Sheet1")
        self.assertEqual(len(df), 3)
        self.assertEqual(list(df.columns), ["ID", "Name", "Age"])

    def test_compare_by_sequence_identical(self):
        """测试序列比对：完全相同"""
        df1 = ExcelDiffer.read_excel_raw(self.file1, "Sheet1")
        all_cols, results = ExcelDiffer.compare_dataframes(df1, df1)
        
        for left, right, info in results:
            self.assertEqual(info['status'], 'equal')
            self.assertEqual(left, right)

    def test_compare_by_sequence_diff(self):
        """测试序列比对：存在差异"""
        df1 = ExcelDiffer.read_excel_raw(self.file1, "Sheet1")
        df2 = ExcelDiffer.read_excel_raw(self.file2, "Sheet1")
        # 序列比对不关注 ID，只按行号对齐
        all_cols, results = ExcelDiffer.compare_dataframes(df1, df2)
        
        # 第三行应该被标记为修改 (Charlie -> David)
        self.assertEqual(results[2][2]['status'], 'modified')

    def test_compare_by_keys(self):
        """测试主键比对"""
        df1 = ExcelDiffer.read_excel_raw(self.file1, "Sheet1")
        df2 = ExcelDiffer.read_excel_raw(self.file2, "Sheet1")
        
        all_cols, results = ExcelDiffer.compare_dataframes(df1, df2, key_columns=["ID"])
        
        # 预期结果：
        # ID=1: equal
        # ID=2: modified (Age: 30 -> 31)
        # ID=3: deleted
        # ID=4: added
        
        status_map = {r[0][1]: r[2]['status'] for r in results if r[0][1] != ""} # r[0][1] 是 ID 列
        # 注意：added 行的 left_row 第一列是空的，ID 在右侧
        added_id = [r[1][1] for r in results if r[2]['status'] == 'added'][0]
        
        self.assertEqual(status_map.get('1'), 'equal')
        self.assertEqual(status_map.get('2'), 'modified')
        self.assertEqual(status_map.get('3'), 'deleted')
        self.assertEqual(str(added_id), '4')

    def test_missing_key_columns(self):
        """测试缺少主键列时的异常处理"""
        df1 = ExcelDiffer.read_excel_raw(self.file1, "Sheet1")
        with self.assertRaises(Exception) as cm:
            ExcelDiffer.compare_dataframes(df1, df1, key_columns=["NonExistent"])
        self.assertIn("缺少关键列", str(cm.exception))

    def test_empty_dataframe(self):
        """测试空数据比对"""
        df_empty = ExcelDiffer.read_excel_raw(self.file_empty, "Sheet1")
        all_cols, results = ExcelDiffer.compare_dataframes(df_empty, df_empty)
        self.assertEqual(len(results), 0)

if __name__ == '__main__':
    unittest.main()
