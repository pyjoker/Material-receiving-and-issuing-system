"""主程式 - 展示如何使用 ExcelReader"""

from excel_reader import ExcelReader
from pathlib import Path


def main():
    # Excel 檔案路徑（請修改為您的實際檔案路徑）
    file_path = Path("C:\\Users\\YA\\Downloads\\16P2759A-M0008-003-伸泰-電線電纜(第10期次計價)-114.11.29複製.xls")
    
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
            sheet_name = sheet_names[2]  # 讀取第三個工作表
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
            # 讀取 C、T、U 欄
            df4 = reader.read_sheet(sheet_name, usecols="C,T,U")
            print(f"\n讀取 C、T、U 欄的資料（刪除前）:")
            print(f"總列數: {len(df4)}")
            print(df4.head())
            
            # 刪除 T 欄為空值的列
            df4_cleaned = df4.dropna(subset=[df4.columns[1]])  # T 欄是第二欄（索引 1）
            df4_cleaned =  df4_cleaned.iloc[2:]
            print(f"\n刪除前兩列和 T 欄空值後的資料:")
            print(f"總列數: {len(df4_cleaned)}")
            print(df4_cleaned.head(100))
            print()
            
    except FileNotFoundError as e:
        print(f"錯誤: {e}")
        print("請確保 Excel 檔案存在於指定路徑")
    except ValueError as e:
        print(f"錯誤: {e}")
    except Exception as e:
        print(f"發生未預期的錯誤: {e}")


if __name__ == "__main__":
    main()