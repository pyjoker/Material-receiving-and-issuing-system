"""主程式 - 展示如何使用 ExcelReader"""

from excel_reader import ExcelReader
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys

# 將專案根目錄加入路徑以便導入 config
sys.path.append(str(Path(__file__).parent.parent))
from config import INPUT_FILE_PATH, OUTPUT_DIR, DEFAULT_SHEET_INDEX, COLUMNS_TO_READ, COLUMN_RENAME_MAP


def save_to_excel(df, output_path, sheet_name='Sheet1'):
    """
    儲存 DataFrame 到 Excel 檔案
    
    Args:
        df: 要儲存的 DataFrame
        output_path: 輸出檔案路徑（字串或 Path 物件）
        sheet_name: 工作表名稱，預設為 'Sheet1'
    
    Returns:
        bool: 儲存成功返回 True，失敗返回 False
    """
    try:
        output_path = Path(output_path)
        
        # 確保輸出目錄存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 儲存為 Excel 檔案
        with pd.ExcelWriter(output_path, engine='xlwt' if output_path.suffix == '.xls' else 'openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✓ 檔案已成功儲存至: {output_path}")
        return True
        
    except Exception as e:
        print(f"\n✗ 儲存檔案時發生錯誤: {e}")
        return False


def main():
    # 從 config.py 讀取檔案路徑
    file_path = INPUT_FILE_PATH
    
    try:
        # 使用 with 語句自動管理檔案
        with ExcelReader(file_path) as reader:
            # 1. 列出所有工作表
            print("=" * 50)
            print("可用的工作表:")
            print("=" * 50)
            sheet_names = reader.get_sheet_names()
            for idx, name in enumerate(sheet_names):
                print(f"{idx}: {name}")
            print()
            
            # 2. 讀取特定工作表（使用名稱）
            print("=" * 50)
            print("讀取工作表（使用名稱）:")
            print("=" * 50)
            sheet_name = sheet_names[DEFAULT_SHEET_INDEX]  # 從 config 讀取工作表索引
            df1 = reader.read_sheet(sheet_name)
            print(f"\n工作表 '{sheet_name}' 的前 5列:")
            print(df1.head())
            print(f"\n形狀: {df1.shape}")
            print()
            
            # 3. 讀取特定工作表（使用索引）
            print("=" * 50)
            print("讀取工作表（使用索引）:")
            print("=" * 50)
            df2 = reader.read_sheet(0)  # 使用索引讀取第一個工作表
            print(f"\n工作表索引 0 的前 5 列:")
            print(df2.head())
            print()
            
            # 4. 讀取工作表並進行預處理
            print("=" * 50)
            print("讀取工作表並預處理:")
            print("=" * 50)
            df3 = reader.read_sheet_with_preprocessing(
                sheet_name,
                drop_empty_rows=True,
                drop_empty_cols=True,
                fill_na=0  # 將空值填充為 0
            )
            print(f"\n預處理後的資料（前 5 列）:")
            print(df3.head())
            print(f"\n形狀: {df3.shape}")
            print()
            
            # 5. 取得工作表資訊
            print("=" * 50)
            print("工作表資訊:")
            print("=" * 50)
            info = reader.get_sheet_info(sheet_name)
            for key, value in info.items():
                print(f"{key}: {value}")
            print()
            
           # 6. 讀取特定欄位
            print("=" * 50)
            print("讀取特定欄位:")
            print("=" * 50)
            # 從 config 讀取要處理的欄位
            df4 = reader.read_sheet(sheet_name, usecols=COLUMNS_TO_READ)
            print(f"\n讀取 C、T、U 欄的資料（刪除前）:")
            print(f"總列數: {len(df4)}")
            print(df4.head())
            
            # 刪除 T 欄為空值的列
            df4_cleaned = df4.dropna(subset=[df4.columns[1]])  # T 欄是第二欄（索引 1）
            df4_cleaned =  df4_cleaned.iloc[2:]
            # 刪除最後一列
            df4_cleaned = df4_cleaned.iloc[:-1]
            
            # 從 config 讀取欄位名稱對應
            df4_cleaned.columns = [COLUMN_RENAME_MAP[i] for i in range(len(df4_cleaned.columns))]
            
            # 排序項次（格式為 "數字-數字"）
            # 將項次拆分為兩個數字欄位進行排序
            df4_cleaned['sort_key1'] = df4_cleaned['項次'].astype(str).str.split('-').str[0].astype(float)
            df4_cleaned['sort_key2'] = df4_cleaned['項次'].astype(str).str.split('-').str[1].astype(float)
            df4_cleaned = df4_cleaned.sort_values(by=['sort_key1', 'sort_key2'])
            # 移除排序用的輔助欄位
            df4_cleaned = df4_cleaned.drop(columns=['sort_key1', 'sort_key2'])
            # 重置索引
            df4_cleaned = df4_cleaned.reset_index(drop=True)
            
            print(f"\n刪除前兩列和 T 欄空值後的資料（已排序）:")
            print(f"總列數: {len(df4_cleaned)}")
            print(df4_cleaned.head(100))
            print()
            
            # 7. 儲存處理後的資料到新的 Excel 檔案
            print("=" * 50)
            print("儲存處理後的資料:")
            print("=" * 50)
            
            # 生成輸出檔案名稱（加上時間戳記）
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # 從 config 讀取輸出目錄
            output_path = OUTPUT_DIR / f"processed_data_{timestamp}.xlsx"
            
            # 儲存檔案
            save_to_excel(df4_cleaned, output_path, sheet_name='處理後資料')
            
    except FileNotFoundError as e:
        print(f"錯誤: {e}")
        print("請確保 Excel 檔案存在於指定路徑")
    except ValueError as e:
        print(f"錯誤: {e}")
    except Exception as e:
        print(f"發生未預期的錯誤: {e}")


if __name__ == "__main__":
    main()