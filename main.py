"""
主程式：自動產生稽核報表並寄送郵件
執行順序：
1. 呼叫 automate.generate_reports() - 從資料庫產生 Excel 報表
2. 呼叫 send_email.send_email() - 自動寄送郵件
"""

from automate import generate_reports
from send_email import send_email

def main():
    """主流程"""
    print("=" * 60)
    print("股務系統月報自動化流程")
    print("=" * 60)

    # ========== 步驟 1: 產生 Excel 報表 ==========
    print("\n【步驟 1/2】開始產生 Excel 報表...")
    print("-" * 60)

    try:
        generate_reports()
        print("\n✓ Excel 報表產生完成！")
    except Exception as e:
        print(f"\n✗ 錯誤：產生報表失敗")
        print(f"錯誤訊息：{e}")
        return

    # ========== 步驟 2: 寄送郵件 ==========
    print("\n" + "=" * 60)
    print("【步驟 2/2】開始寄送郵件...")
    print("-" * 60)

    try:
        result = send_email()
        if result:
            print("\n✓ 郵件處理完成！")
        else:
            print("\n⚠ 郵件處理失敗")
    except Exception as e:
        print(f"\n✗ 錯誤：寄送郵件失敗")
        print(f"錯誤訊息：{e}")
        return

    # ========== 完成 ==========
    print("\n" + "=" * 60)
    print("✓ 所有流程執行完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
