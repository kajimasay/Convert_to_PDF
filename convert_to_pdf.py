import os
import win32com.client
from tkinter import Tk
from tkinter.filedialog import askdirectory

# 📁 GUIで対象のExcelファイルがあるフォルダを選択
Tk().withdraw()
input_folder = askdirectory(title="PDF変換したいExcelファイルのフォルダを選択してください")

if not input_folder:
    print("❌ フォルダが選択されませんでした。終了します。")
    exit()

# 📁 選択されたフォルダ名を取得（末尾フォルダ名）
base_folder_name = os.path.basename(input_folder)
output_folder = os.path.join(os.path.dirname(input_folder), f"{base_folder_name}_pdf_output")
os.makedirs(output_folder, exist_ok=True)

# 📦 Excelアプリ起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# 🔁 .xlsxファイルを一括PDF化
for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx") and not filename.startswith("~$"):
        full_path = os.path.join(input_folder, filename)
        pdf_name = filename.replace(".xlsx", ".pdf")
        pdf_path = os.path.join(output_folder, pdf_name)

        try:
            wb = excel.Workbooks.Open(full_path, ReadOnly=False)
            wb.ExportAsFixedFormat(0, pdf_path)
            wb.Close(False)
            print(f"✅ PDF出力成功: {pdf_name}")
        except Exception as e:
            print(f"❌ PDF出力失敗: {pdf_name} → {e}")

# Excel終了
excel.Quit()
print(f"\n✅ すべてのExcelファイルをPDFに変換しました！出力先: {output_folder}")
