import os
import win32com.client

# Excelファイルが保存されているフォルダ（変更してください）
input_folder = r"G:\マイドライブ\Cell Vision Global Limited\発注からインボイス変換\invoice_app\OCS_invoice_250509"
# PDFを保存するフォルダ（任意で変更可能）
output_folder = r"G:\マイドライブ\Cell Vision Global Limited\発注からインボイス変換\invoice_app\pdf_output"

os.makedirs(output_folder, exist_ok=True)

# Excel アプリケーションを起動（バックグラウンドで）
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx") and not filename.startswith("~$"):
        full_path = os.path.join(input_folder, filename)
        wb = excel.Workbooks.Open(full_path)

        pdf_name = filename.replace(".xlsx", ".pdf")
        pdf_path = os.path.join(output_folder, pdf_name)

        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)

excel.Quit()
print("✅ すべてのExcelファイルをPDFに変換しました！")
