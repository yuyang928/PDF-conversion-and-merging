import win32com.client
from PyPDF2 import PdfMerger
import os

def word_to_pdf(word_path, pdf_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(word_path)
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
    doc.Close()
    word.Quit()

def excel_to_pdf(
    excel_path,
    pdf_path,
    sheets=None   # None = 所有 sheet；或者 ["Sheet1", "Sheet2"]
):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(excel_path))

    for ws in wb.Worksheets:
        if sheets and ws.Name not in sheets:
            continue

        ps = ws.PageSetup

        # ✅ 横向打印
        ps.Orientation = 2        # 2 = Landscape

        # ✅ 关键：关闭 Zoom，启用 FitToPages
        ps.Zoom = False

        # ✅ 一页宽 + 一页高（最适合超宽表）
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1

        # ✅ 页面边距（单位：英寸）
        ps.LeftMargin = excel.InchesToPoints(0.25)
        ps.RightMargin = excel.InchesToPoints(0.25)
        ps.TopMargin = excel.InchesToPoints(0.5)
        ps.BottomMargin = excel.InchesToPoints(0.5)

        # ✅ 居中（可选）
        ps.CenterHorizontally = True
        ps.CenterVertically = False

    # ✅ 导出 PDF
    wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # 0 = PDF

    wb.Close(False)
    excel.Quit()

def merge_pdfs(pdf_list, output_pdf):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()

if __name__ == "__main__":
    files = [
        "yuyang/files/lab4d.docx",
        "yuyang/files/lab4x.xlsx",
        "yuyang/files/exercise_3.docx",
    ]

    temp_pdfs = []

    for f in files:
        pdf = os.path.splitext(f)[0] + ".pdf"
        if f.lower().endswith(".docx"):
            word_to_pdf(os.path.abspath(f), os.path.abspath(pdf))
        elif f.lower().endswith(".xlsx"):
            excel_to_pdf(os.path.abspath(f), os.path.abspath(pdf))
        temp_pdfs.append(pdf)

    merge_pdfs(temp_pdfs, "final_merged.pdf")
    print("✅ 合并完成：final_merged.pdf")