"""Excel Reader Module for Material Receiving and Issuing System"""

import pandas as pd
from pathlib import Path
from typing import Union, List


class ExcelReader:
    """讀取 Excel 檔案的類別，支援多工作表檔案"""
    
    def __init__(self, file_path: Union[str, Path]):
        """
        初始化 ExcelReader
        
        Args:
            file_path: Excel 檔案路徑
        
        Raises:
            FileNotFoundError: 如果檔案不存在
            ValueError: 如果檔案格式不正確
        """
        self.file_path = Path(file_path)
        
        if not self.file_path.exists():
            raise FileNotFoundError(f"找不到檔案: {self.file_path}")
        
        if self.file_path.suffix not in ['.xlsx', '.xls']:
            raise ValueError(f"不支援的檔案格式: {self.file_path.suffix}，請使用 .xlsx 或 .xls")
        
        self._excel_file = pd.ExcelFile(self.file_path)
    
    def get_sheet_names(self) -> List[str]:
        """
        取得所有工作表名稱
        
        Returns:
            工作表名稱列表
        """
        return self._excel_file.sheet_names
    
    def read_sheet(self, 
                   sheet: Union[str, int], 
                   header: int = 0,
                   skiprows: int = None,
                   usecols: Union[str, List] = None) -> pd.DataFrame:
        """
        讀取指定的工作表
        
        Args:
            sheet: 工作表名稱（字串）或索引（數字，從 0 開始）
            header: 標題列位置，預設為 0（第一列）
            skiprows: 跳過前幾列
            usecols: 要讀取的欄位，可以是欄位名稱列表或 Excel 欄位範圍（如 "A:C"）
        
        Returns:
            包含工作表資料的 DataFrame
        
        Raises:
            ValueError: 如果工作表不存在
        """
        # 驗證工作表是否存在
        if isinstance(sheet, str) and sheet not in self.get_sheet_names():
            raise ValueError(f"工作表 '{sheet}' 不存在。可用的工作表: {', '.join(self.get_sheet_names())}")
        
        if isinstance(sheet, int) and (sheet < 0 or sheet >= len(self.get_sheet_names())):
            raise ValueError(f"工作表索引 {sheet} 超出範圍。有效範圍: 0-{len(self.get_sheet_names())-1}")
        
        # 讀取工作表
        try:
            df = pd.read_excel(
                self._excel_file,
                sheet_name=sheet,
                header=header,
                skiprows=skiprows,
                usecols=usecols
            )
            return df
        except Exception as e:
            raise RuntimeError(f"讀取工作表時發生錯誤: {str(e)}")
    
    def read_sheet_with_preprocessing(self,
                                     sheet: Union[str, int],
                                     drop_empty_rows: bool = True,
                                     drop_empty_cols: bool = True,
                                     fill_na: Union[str, int, float] = None) -> pd.DataFrame:
        """
        讀取工作表並進行預處理
        
        Args:
            sheet: 工作表名稱或索引
            drop_empty_rows: 是否刪除空白列
            drop_empty_cols: 是否刪除空白欄
            fill_na: 填充空值的值（None 表示不填充）
        
        Returns:
            預處理後的 DataFrame
        """
        df = self.read_sheet(sheet)
        
        # 刪除空白列
        if drop_empty_rows:
            df = df.dropna(how='all')
        
        # 刪除空白欄
        if drop_empty_cols:
            df = df.dropna(axis=1, how='all')
        
        # 填充空值
        if fill_na is not None:
            df = df.fillna(fill_na)
        
        # 重置索引
        df = df.reset_index(drop=True)
        
        return df
    
    def get_sheet_info(self, sheet: Union[str, int]) -> dict:
        """
        取得工作表的基本資訊
        
        Args:
            sheet: 工作表名稱或索引
        
        Returns:
            包含工作表資訊的字典
        """
        df = self.read_sheet(sheet)
        
        return {
            '工作表名稱': sheet if isinstance(sheet, str) else self.get_sheet_names()[sheet],
            '總列數': len(df),
            '總欄數': len(df.columns),
            '欄位名稱': df.columns.tolist(),
            '資料型別': df.dtypes.to_dict(),
            '空值統計': df.isnull().sum().to_dict()
        }
    
    def close(self):
        """關閉 Excel 檔案"""
        self._excel_file.close()
    
    def __enter__(self):
        """支援 with 語句"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """支援 with 語句"""
        self.close()