# write_token_txt.py
import openpyxl

def extract_token():
    try:
        wb = openpyxl.load_workbook("Access_Token.xlsx", data_only=True)
        ws = wb.active
        token = ws["A2"].value
        if not token:
            raise ValueError("A2 is empty or invalid")
        with open("access_token.txt", "w") as f:
            f.write(token)
        print("✅ Token đã được ghi vào access_token.txt")
    except Exception as e:
        print(f"❌ Lỗi khi trích xuất token: {e}")
        exit(1)

if __name__ == "__main__":
    extract_token()
