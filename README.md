# 股務系統月報自動化

自動從資料庫查詢股務資料並產生加密的 Excel 月報，並透過 Outlook 自動寄送郵件。

## 功能特色

- ✅ 自動從 Oracle 資料庫查詢多間公司資料（101, 103）
- ✅ 產生加密的 Excel 報表（密碼保護）
- ✅ 自動計算上個月起訖日期（民國年格式）
- ✅ 支援多工作表（地址變更、過戶書號、現金股利、股票股利、掛失）
- ✅ 自動透過 Outlook 寄送郵件附件
- ✅ 模組化設計，可單獨執行或整體執行

## 系統需求

- Python 3.8+
- Windows 作業系統（使用 Outlook COM 介面）
- Oracle Instant Client
- Microsoft Outlook

## 安裝步驟

### 1. Clone 專案

```bash
git clone <repository-url>
cd Monthly_report_automate
```

### 2. 安裝套件

```bash
pip install -r requirements.txt
```

### 3. 設定環境變數

複製 `.env.example` 為 `.env` 並填入資料庫連線資訊：

```bash
cp .env.example .env
```

編輯 `.env` 檔案：

```env
# 資料庫連接設定
DB_USER=your_username
DB_PASSWORD=your_password
DB_HOST=10.37.34.74
DB_PORT=1523
DB_SERVICE=SKMS01

# Oracle Client 路徑
ORACLE_CLIENT_PATH=C:\instantclient_21_19
```

### 4. 設定 SQL 查詢

將 SQL 查詢檔案放在 `sql/` 資料夾下：
- `changeAddress.sql` - 地址變更查詢
- `transferNumber.sql` - 過戶書號查詢
- `cashDividend.sql` - 現金股利查詢
- `stockDividend.sql` - 股票股利查詢
- `lost.sql` - 掛失查詢

### 5. 設定郵件資訊

編輯 `send_email.py` 中的郵件設定（第 11-37 行）：

```python
recipient_email = "recipient@example.com"  # 收件者
cc_email = ""  # 副本（選填）
sender_email = "your.email@company.com"  # 寄件者
subject_template = "RE: 股務系統每月月報"  # 主旨
send_mode = "display"  # 寄送模式
```

## 使用方式

### 執行完整流程

執行主程式，依序產生報表並寄送郵件：

```bash
python main.py
```

### 單獨執行功能

**只產生報表：**
```bash
python automate.py
```

**只寄送郵件：**
```bash
python send_email.py
```

## 檔案說明

```
Monthly_report_automate/
├── main.py                  # 主程式（執行完整流程）
├── automate.py              # 產生 Excel 報表
├── send_email.py            # 寄送郵件
├── check_outlook_accounts.py # 檢查 Outlook 帳號工具
├── .env                     # 環境變數（不會被 git 追蹤）
├── .env.example             # 環境變數範例
├── .gitignore               # Git 忽略檔案
├── requirements.txt         # Python 套件清單
├── README.md                # 說明文件
└── sql/                     # SQL 查詢檔案資料夾
    ├── changeAddress.sql
    ├── transferNumber.sql
    ├── cashDividend.sql
    ├── stockDividend.sql
    └── lost.sql
```

## 產出檔案

執行後會產生兩個加密的 Excel 檔案：

- `公司代號101_稽核資料_1140901~1140930.xlsx`
- `公司代號103_稽核資料_1140901~1140930.xlsx`

**開啟密碼**: `8502`

## 郵件寄送模式

編輯 `send_email.py` 中的 `send_mode` 設定：

- **`"display"`** - 開啟郵件視窗讓你檢查後手動寄出（推薦）
- **`"send"`** - 直接寄出
- **`"draft"`** - 存到草稿匣

## 自訂設定

### 修改公司代號

編輯 `automate.py` 第 21 行：

```python
company_codes = ['101', '103']  # 新增或修改公司代號
```

### 修改欄位對應

編輯 `automate.py` 第 42-48 行的 `column_mappings` 字典：

```python
column_mappings = {
    "現金股利": ["公司代號", "年度", "戶號", "實發股利金額", "領取日", "領取編號"],
    # ... 其他對應
}
```

### 修改 Excel 加密密碼

編輯 `automate.py` 第 133 行：

```python
password = "8502"  # 修改為你的密碼
```

## 常見問題

### Q: 為什麼寄信失敗？
A: 確認 Outlook 已開啟，並檢查 `sender_email` 設定是否正確。

### Q: 如何查看我的 Outlook 帳號？
A: 執行 `python check_outlook_accounts.py` 查看所有可用帳號。

### Q: Excel 檔案無法開啟？
A: 輸入密碼 `8502` 即可開啟。

### Q: 資料庫連線失敗？
A: 檢查 `.env` 檔案中的資料庫設定是否正確，並確認 Oracle Client 已安裝。

## 技術細節

### 使用套件

- **oracledb** - Oracle 資料庫連線
- **pandas** - 資料處理
- **openpyxl** - Excel 檔案操作
- **msoffcrypto-tool** - Excel 檔案加密
- **pywin32** - Outlook COM 介面
- **python-dotenv** - 環境變數管理

### 日期計算

- 自動計算執行日期的「上個月」起訖日期
- 轉換為民國年格式（例如：1140901~1140930）

### SQL 參數化

- 使用正則表達式動態替換 SQL 中的公司代號
- 支援不同格式：`c_comp_co='101'` 或 `c_comp_co = '101'`

## License

MIT

## 作者

Generated with Claude Code
