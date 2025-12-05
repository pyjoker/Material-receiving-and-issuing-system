"""範例：使用 WebFormFiller 自動填寫網頁表單"""

from excel_reader import ExcelReader
from web_form_filler import WebFormFiller, fill_web_form_from_dataframe
from pathlib import Path
import pandas as pd
from datetime import datetime
import sys

# 將專案根目錄加入路徑以便導入 config
sys.path.append(str(Path(__file__).parent.parent))
from config import INPUT_FILE_PATH, DEFAULT_SHEET_INDEX, COLUMNS_TO_READ, COLUMN_RENAME_MAP, PROCESSED_FILE_PATH


def example_basic_usage():
    """基本使用範例"""
    
    # 1. 讀取 Excel 檔案並處理資料
    print("=" * 50)
    print("步驟 1: 讀取並處理 Excel 資料")
    print("=" * 50)
    
    # 從 config.py 讀取檔案路徑
    file_path = INPUT_FILE_PATH
    
    with ExcelReader(file_path) as reader:
        sheet_names = reader.get_sheet_names()
        sheet_name = sheet_names[DEFAULT_SHEET_INDEX]  # 從 config 讀取工作表索引
        
        # 從 config 讀取要處理的欄位
        df = reader.read_sheet(sheet_name, usecols=COLUMNS_TO_READ)
        
        # 資料處理
        df = df.dropna(subset=[df.columns[1]])  # 刪除 T 欄空值
        df = df.iloc[2:]  # 刪除前兩列
        df = df.iloc[:-1]  # 刪除最後一列
        
        # 從 config 讀取欄位名稱對應
        df.columns = [COLUMN_RENAME_MAP[i] for i in range(len(df.columns))]
        
        # 排序
        df['sort_key1'] = df['項次'].astype(str).str.split('-').str[0].astype(float)
        df['sort_key2'] = df['項次'].astype(str).str.split('-').str[1].astype(float)
        df = df.sort_values(by=['sort_key1', 'sort_key2'])
        df = df.drop(columns=['sort_key1', 'sort_key2'])
        df = df.reset_index(drop=True)
        
        print(f"✓ 已載入 {len(df)} 筆資料")
        print("\n前 5 筆資料:")
        print(df.head())
    
    # 2. 填寫網頁表單
    print("\n" + "=" * 50)
    print("步驟 2: 自動填寫網頁表單")
    print("=" * 50)
    
    # 請替換為實際的網頁 URL
    url = "https://ctcieip.ctci.com/pp_mrs/PP_MRS_3010.aspx?ParentAPPL=F:$VSTS02_CCC$PMS$&HostUrl=ctcieip.ctci.com"
    
    # 方法 1: 使用便捷函式（推薦）
    results = fill_web_form_from_dataframe(
        df=df,
        url=url,
        headless=False,  # 設為 True 可隱藏瀏覽器視窗
        wait_time=10,    # 等待頁面載入 10 秒
        delay=0.5        # 每筆資料間隔 0.5 秒
    )
    
    print("\n處理結果:", results)


def example_manual_control():
    """手動控制範例（進階用法）"""
    
    # 準備資料
    print("=" * 50)
    print("步驟 1: 選擇資料來源")
    print("=" * 50)
    print("1. 未處理的 Excel 原始資料")
    print("2. 已處理的 Excel 資料（processed_data_*.xlsx）")
    
    data_choice = input("\n請選擇資料來源 (1/2): ").strip()
    
    if data_choice == "1":
        # 未處理資料路徑 - 從原始 Excel 讀取並處理
        print("\n讀取未處理的資料...")
        file_path = INPUT_FILE_PATH
        
        with ExcelReader(file_path) as reader:
            sheet_names = reader.get_sheet_names()
            sheet_name = sheet_names[DEFAULT_SHEET_INDEX]  # 從 config 讀取工作表索引
            
            # 從 config 讀取要處理的欄位
            df = reader.read_sheet(sheet_name, usecols=COLUMNS_TO_READ)
            
            # 資料處理
            df = df.dropna(subset=[df.columns[1]])  # 刪除 T 欄空值
            df = df.iloc[2:]  # 刪除前兩列
            df = df.iloc[:-1]  # 刪除最後一列
            
            # 從 config 讀取欄位名稱對應
            df.columns = [COLUMN_RENAME_MAP[i] for i in range(len(df.columns))]
            
            # 排序
            df['sort_key1'] = df['項次'].astype(str).str.split('-').str[0].astype(float)
            df['sort_key2'] = df['項次'].astype(str).str.split('-').str[1].astype(float)
            df = df.sort_values(by=['sort_key1', 'sort_key2'])
            df = df.drop(columns=['sort_key1', 'sort_key2'])
            df = df.reset_index(drop=True)
            
            print(f"✓ 已載入 {len(df)} 筆資料")
            print("\n前 5 筆資料:")
            print(df.head())
    
    elif data_choice == "2":
        # 已處理資料路徑 - 直接從 config 讀取
        print("\n讀取已處理的資料...")
        processed_file_path = PROCESSED_FILE_PATH
        
        if not processed_file_path.exists():
            print(f"✗ 檔案不存在: {processed_file_path}")
            print("提示: 請在 config.py 中設定正確的 PROCESSED_FILE_PATH")
            return
        
        # 直接讀取為 DataFrame
        df = pd.read_excel(processed_file_path)
        print(f"✓ 已從 {processed_file_path.name} 載入 {len(df)} 筆資料")
        print("\n前 5 筆資料:")
        print(df.head())
    
    else:
        print("無效的選項")
        return
    
    # 建立 WebFormFiller 實例
    filler = WebFormFiller(headless=False)
    
    try:
        # 啟動瀏覽器
        filler.start_browser()
        
        # 開啟網頁
        url = "https://ctcieip.ctci.com/pp_mrs/PP_MRS_3010.aspx?ParentAPPL=F:$VSTS02_CCC$PMS$&HostUrl=ctcieip.ctci.com"
        filler.open_url(url, wait_time=10)
        
        # 如果需要登入或其他操作，可以在這裡手動處理
        # 例如：
        input("請手動登入網站，完成後按 Enter 繼續...")
        
        # 開始計時
        start_time = datetime.now()
        print(f"\n開始時間: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 處理資料
        results = filler.process_dataframe(df, delay=0.1)
        
        # 結束計時
        end_time = datetime.now()
        elapsed_time = end_time - start_time
        print(f"\n結束時間: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"總耗時: {elapsed_time.total_seconds():.2f} 秒 ({elapsed_time})")
        
        # 在這裡可以做其他操作
        # 例如點擊儲存按鈕等
        # save_button = filler.driver.find_element(By.ID, "btnSave")
        # save_button.click()
        
        print("\n處理結果:", results)
        
        # 讓瀏覽器保持開啟，方便檢查結果
        input("\n按 Enter 關閉瀏覽器...")
        
    finally:
        # 關閉瀏覽器
        filler.close_browser()


def example_with_login():
    """包含登入的完整範例"""
    
    # 讀取資料
    file_path = Path("C:\\Users\\YA\\Downloads\\16P2759A-M0008-003-伸泰-電線電纜(第10期次計價)-114.11.29複製.xls")
    
    with ExcelReader(file_path) as reader:
        sheet_name = reader.get_sheet_names()[2]
        df = reader.read_sheet(sheet_name, usecols="C,T,U")
        df = df.dropna(subset=[df.columns[1]]).iloc[2:-1]
        df.columns = ['項次', '數量', '複價']
        
        # 排序
        df['sort_key1'] = df['項次'].astype(str).str.split('-').str[0].astype(float)
        df['sort_key2'] = df['項次'].astype(str).str.split('-').str[1].astype(float)
        df = df.sort_values(by=['sort_key1', 'sort_key2']).drop(columns=['sort_key1', 'sort_key2']).reset_index(drop=True)
    
    # 建立填寫器
    with WebFormFiller(headless=False) as filler:
        # 開啟登入頁面
        login_url = "https://your-login-page.com"
        filler.open_url(login_url, wait_time=5)
        
        # 等待手動登入（或自動填寫登入資訊）
        print("\n請在瀏覽器中完成登入操作...")
        input("登入完成後，按 Enter 繼續...")
        
        # 導航到目標頁面
        target_url = "https://your-target-page.com"
        filler.open_url(target_url, wait_time=10)
        
        # 處理資料
        results = filler.process_dataframe(df, delay=0.5)
        
        print("\n處理完成！")
        input("按 Enter 關閉瀏覽器...")


if __name__ == "__main__":
    example_manual_control()
    # print("網頁表單自動填寫範例\n")
    # print("請選擇要執行的範例:")
    # print("1. 基本使用範例")
    # print("2. 手動控制範例（進階）")
    # print("3. 包含登入的完整範例")
    
    # choice = input("\n請輸入選項 (1/2/3): ").strip()
    
    # if choice == "1":
    #     example_basic_usage()
    # elif choice == "2":
    #     example_manual_control()
    # elif choice == "3":
    #     example_with_login()
    # else:
    #     print("無效的選項")
