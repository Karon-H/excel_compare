from abc import ABC, abstractmethod
import pandas as pd

class BaseDiffer(ABC):
    """
    比对引擎抽象基类
    定义了所有比对引擎必须实现的标准接口
    """

    @abstractmethod
    def get_containers(self, filepath):
        """
        获取文件中的数据容器列表
        例如：Excel 的 Sheet 列表，数据库的表列表
        """
        pass

    @abstractmethod
    def read_data(self, filepath, container_name, **kwargs):
        """
        从指定容器读取数据并返回为 pandas.DataFrame
        """
        pass

    @abstractmethod
    def compare(self, df1, df2, key_columns=None, **kwargs):
        """
        比对两个 DataFrame 并返回列名列表和比对结果列表
        """
        pass

    @abstractmethod
    def export(self, filepath, columns, results, key_columns=None, **kwargs):
        """
        导出比对结果到文件
        """
        pass

    @abstractmethod
    def compare_all(self, file1, file2, key_columns=None, progress_callback=None, **kwargs):
        """
        自动匹配并比对两个文件中的所有容器，返回摘要列表
        """
        pass
