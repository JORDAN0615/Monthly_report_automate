import win32com.client
import os
from datetime import datetime
from calendar import monthrange

def send_email():
    """寄送稽核報表郵件的主函數"""

    # ========== 設定區 ==========
    # 收件者 Email（多個收件者用分號分隔，例如："email1@example.com; email2@example.com"）
    recipient_email = "jordan00661155@gmail.com"

    # 副本（選填）
    cc_email = ""

    # 寄件者帳號（選填，留空則使用 Outlook 預設帳號）
    # 填入你的完整 Email，例如："your.name@company.com"
    sender_email = "Jordan_Tseng@wistron.com"

    # 郵件主旨
    subject_template = "RE: 股務系統每月月報"

    # 郵件內容
    email_body = """Hi Nina,

        附件為股務系統月報，請查收。

        如有任何問題，請隨時聯繫。

        謝謝！
        """

    # 郵件處理模式
    # "display" = 開啟郵件視窗讓你檢查後手動寄出
    # "send" = 直接寄出
    # "draft" = 存到草稿匣
    send_mode = "display"

    # ========== 計算日期範圍 ==========
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

    # ========== 尋找要附加的檔案 ==========
    attachment_files = []
    company_codes = ['101', '103']

    for company_code in company_codes:
        filename = f"公司代號{company_code}_稽核資料_{date_range}.xlsx"
        file_path = os.path.abspath(filename)

        if os.path.exists(file_path):
            attachment_files.append(file_path)
            print(f"✓ 找到檔案：{filename}")
        else:
            print(f"✗ 找不到檔案：{filename}")

    if not attachment_files:
        print("\n錯誤：沒有找到任何附件檔案！")
        return False

    # ========== 建立 Outlook 郵件 ==========
    print(f"\n開始建立 Outlook 郵件...")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = MailItem

    # 設定寄件者（如果有指定）
    if sender_email:
        # 尋找指定的帳號
        for account in outlook.Session.Accounts:
            if account.SmtpAddress.lower() == sender_email.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                print(f"✓ 使用寄件者帳號：{sender_email}")
                break
        else:
            print(f"⚠ 找不到帳號 {sender_email}，將使用預設帳號")

    # 設定收件者
    mail.To = recipient_email

    # 設定副本（如果有）
    if cc_email:
        mail.CC = cc_email

    # 設定主旨
    mail.Subject = subject_template.format(date_range=date_range)

    # 設定郵件內容
    mail.Body = email_body

    # 附加檔案
    for file_path in attachment_files:
        mail.Attachments.Add(file_path)
        print(f"✓ 已附加：{os.path.basename(file_path)}")

    # 先儲存郵件資訊（因為寄出後物件會被釋放）
    email_subject = mail.Subject
    attachment_count = len(attachment_files)

    # 根據模式處理郵件
    if send_mode == "send":
        mail.Send()
        print(f"\n✓ 郵件已寄出！")
    elif send_mode == "display":
        mail.Display()
        print(f"\n✓ 郵件視窗已開啟，請檢查後按「傳送」寄出！")
    else:  # draft
        mail.Save()
        print(f"\n✓ 郵件已存到草稿匣，請到 Outlook 檢查後再寄出！")

    print(f"收件者：{recipient_email}")
    print(f"主旨：{email_subject}")
    print(f"附件數量：{attachment_count} 個")

    return True


# 如果直接執行此檔案
if __name__ == "__main__":
    send_email()
