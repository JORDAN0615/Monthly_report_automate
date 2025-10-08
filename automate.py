import oracledb
import pandas as pd
import os
import re
from datetime import datetime
from calendar import monthrange
from dotenv import load_dotenv

def generate_reports():
    """產生稽核報表的主函數"""

    # 載入 .env 檔案
    load_dotenv()

    # 從環境變數讀取資料庫連接資訊
    oracle_client_path = os.getenv('ORACLE_CLIENT_PATH')
    db_user = os.getenv('DB_USER')
    db_password = os.getenv('DB_PASSWORD')
    db_host = os.getenv('DB_HOST')
    db_port = os.getenv('DB_PORT')
    db_service = os.getenv('DB_SERVICE')

    oracledb.init_oracle_client(lib_dir=oracle_client_path)

    # 資料庫連接
    dsn = f"{db_host}:{db_port}/{db_service}"
    connection = oracledb.connect(
        user=db_user,
        password=db_password,
        dsn=dsn
    )

    # 公司代號列表
    company_codes = ['101', '103']

    # 計算上個月的起訖日期（民國年格式）
    now = datetime.now()
    # 計算上個月
    if now.month == 1:
        last_month = 12
        last_year = now.year - 1
    else:
        last_month = now.month - 1
        last_year = now.year

    # 轉換成民國年
    roc_year = last_year - 1911

    # 上個月的最後一天
    last_day = monthrange(last_year, last_month)[1]

    # 格式化日期範圍：1140901~1140930
    date_range = f"{roc_year}{last_month:02d}01~{roc_year}{last_month:02d}{last_day:02d}"

    # 為每個工作表定義不同的欄位名稱
    column_mappings = {
        "地址變更": ["公司代號", "戶號", "戶名", "更址書號", "變更日期"],
        "過戶書號": ["公司代號", "過戶書號", "過戶日期", "出讓人戶號", "受讓人戶號", "過戶股數"],
        "現金股利": ["公司代號", "年度", "戶號", "實發股利金額", "領取日", "領取編號"],
        "股票股利": ["公司代號", "年度", "戶號", "領取股數", "領取日", "領取編號"],
        "掛失": ["公司代號", "戶號", "掛失日期", "掛失書號"],
    }

    # SQL 檔案名稱及對應的工作表名稱
    sql_files = [
        "sql/changeAddress.sql",
        "sql/transferNumber.sql",
        "sql/cashDividend.sql",
        "sql/stockDividend.sql",
        "sql/lost.sql"
    ]

    sheet_names = [
        "地址變更",
        "過戶書號",
        "現金股利",
        "股票股利",
        "掛失"
    ]

    # 讀取 SQL 檔案內容成列表
    sql_queries = []
    for file in sql_files:
        if os.path.exists(file):
            with open(file, 'r', encoding='utf-8') as f:
                sql = f.read().strip()  # 讀取並去除多餘空白
            sql_queries.append(sql)
            print(f"已讀取 {file}")
        else:
            print(f"錯誤：找不到檔案 {file}")
            sql_queries.append("")  # 空 SQL，避免崩潰

    # 為每間公司產生 Excel 檔案
    for company_code in company_codes:
        # 檔案名稱格式：公司代號103_稽核資料_1140901~1140930.xlsx
        output_file = f"公司代號{company_code}_稽核資料_{date_range}.xlsx"
        print(f"\n開始處理公司代號 {company_code}...")

        # 開啟 Excel 寫入器
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            for i, sql in enumerate(sql_queries, 1):
                if not sql:  # 跳過空 SQL
                    print(f"  Query {i} 跳過：無 SQL")
                    continue
                try:
                    # 動態替換 SQL 中的公司代號（支援不同空格格式）
                    modified_sql = re.sub(
                        r"c_comp_co\s*=\s*'101'",  # 匹配 c_comp_co='101' 或 c_comp_co = '101'
                        f"c_comp_co='{company_code}'",
                        sql,
                        flags=re.IGNORECASE
                    )

                    # 執行 SQL 並讀取結果成 DataFrame
                    df = pd.read_sql(modified_sql, connection)

                    # 寫入不同工作表（使用自訂名稱）
                    sheet_name = sheet_names[i-1]  # 因為 i 從 1 開始，所以要減 1

                    # 根據工作表名稱設定對應的欄位名稱
                    if sheet_name in column_mappings:
                        df.columns = column_mappings[sheet_name]

                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"  {sheet_name} 完成：{len(df)} 筆資料")
                except Exception as e:
                    print(f"  {sheet_name} 錯誤：{e}")

        print(f"✓ {output_file} 已產生")

    # 關閉連接
    connection.close()
    print("\n所有查詢完成！")


# 如果直接執行此檔案
if __name__ == "__main__":
    generate_reports()