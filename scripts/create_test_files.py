import pandas as pd
import os

# 获取脚本所在目录的父目录，即项目根目录
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS_DIR = os.path.join(BASE_DIR, "docs")

data1 = {
    'ID': [1, 2, 3],
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [30, 25, 35],
    'City': ['NY', 'LA', 'SF']
}

data2 = {
    'ID': [1, 2, 4],
    'Name': ['Alice', 'Bob', 'David'],
    'Age': [31, 25, 28], # Alice age changed
    'City': ['NY', 'LA', 'CHI']
}

df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)

# 确保目录存在
os.makedirs(DOCS_DIR, exist_ok=True)

df1.to_excel(os.path.join(DOCS_DIR, 'test1.xlsx'), index=False)
df2.to_excel(os.path.join(DOCS_DIR, 'test2.xlsx'), index=False)

print(f"测试文件已创建在: {DOCS_DIR}")
