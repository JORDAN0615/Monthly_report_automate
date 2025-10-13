@echo off
chcp 65001 >nul
REM 稽核報表自動產生批次檔
REM 作者: Jordan Tseng
REM 說明: 自動產生上個月的稽核報表

echo ============================================
echo 稽核報表自動產生工具
echo ============================================
echo.

REM 切換到腳本所在目錄
cd /d "%~dp0"

REM 啟動虛擬環境
echo [1/3] 啟動虛擬環境...
call .venv\Scripts\activate.bat

REM 執行 Python 程式
echo [2/3] 執行主程式（產生報表 + 寄送郵件）...
python main.py

REM 檢查執行結果
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo [3/3] 報表產生完成！
    echo ============================================
) else (
    echo.
    echo ============================================
    echo [3/3] 執行過程發生錯誤！錯誤代碼: %ERRORLEVEL%
    echo ============================================
)

echo.
echo 按任意鍵關閉視窗...
pause >nul
