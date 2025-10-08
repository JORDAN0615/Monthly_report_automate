import win32com.client

print("正在檢查 Outlook 帳號...\n")

try:
    outlook = win32com.client.Dispatch("Outlook.Application")

    # 取得所有帳號
    accounts = outlook.Session.Accounts

    print(f"找到 {accounts.Count} 個帳號：\n")

    for i, account in enumerate(accounts, 1):
        print(f"{i}. {account.DisplayName}")
        print(f"   Email: {account.SmtpAddress}")
        print(f"   類型: {account.AccountType}")

        # 檢查是否為預設帳號（通常第一個就是預設）
        if i == 1:
            print(f"   ★ 這是預設帳號")
        print()

    # 顯示預設寄件者
    print("-" * 50)
    print(f"\n預設寄件者會使用：{accounts.Item(1).SmtpAddress}")

except Exception as e:
    print(f"錯誤：{e}")
    print("\n請確認：")
    print("1. Outlook 已經開啟")
    print("2. 已安裝 pywin32 套件")
