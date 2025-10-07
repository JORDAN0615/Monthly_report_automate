import oracledb
import pandas as pd
import os

oracledb.init_oracle_client(lib_dir=r"C:\instantclient_21_19")

# 資料庫連接資訊（用 SQL Developer 的細節替換）
connection = oracledb.connect(
    user="SKMGR",      # 你的 DB 使用者名
    password="SKMGR",  # 你的密碼
    dsn="10.37.34.74:1523/SKMS01"  # e.g., "localhost:1521/XE"
)

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

# 開啟 Excel 寫入器（自動建立 output.xlsx）
with pd.ExcelWriter("monthly_report.xlsx", engine="openpyxl") as writer:
    for i, sql in enumerate(sql_queries, 1):
        if not sql:  # 跳過空 SQL
            print(f"Query {i} 跳過：無 SQL")
            continue
        try:
            # 執行 SQL 並讀取結果成 DataFrame
            df = pd.read_sql(sql, connection)

            # 寫入不同工作表（使用自訂名稱）
            sheet_name = sheet_names[i-1]  # 因為 i 從 1 開始，所以要減 1

            # 根據工作表名稱設定對應的欄位名稱
            if sheet_name in column_mappings:
                df.columns = column_mappings[sheet_name]

            df.to_excel(writer, sheet_name=sheet_name, index=False)


            print(f"Query {i} 完成：{len(df)} 筆資料寫入 {sheet_name}")
        except Exception as e:
            print(f"Query {i} 錯誤：{e}")

# 關閉連接
connection.close()
print("所有查詢完成！檢查 monthly_report.xlsx")