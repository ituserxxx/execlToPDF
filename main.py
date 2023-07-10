import pathlib
import win32com.client
import os
import MegerPDF


def to_pdf(curr_dir_path):
    root_dir = curr_dir_path.strip('\"')
    xlsx_files = []
    for file in os.listdir(root_dir):
        if file.endswith(".xlsx"):
            xlsx_files.append(os.path.join(root_dir, file))
    if len(xlsx_files) == 0:
        print("当前目录没有 .xlsx 的文件")
        return

    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False

    for i in xlsx_files:
        xlsx = pathlib.Path(i)
        xlsx_dir = xlsx.parent
        xlsx_dir = str(xlsx_dir)
        basename = xlsx.stem
        basename = str(basename)
        output_file = xlsx_dir + "/" + basename + ".pdf"
        book = app.Workbooks.Open(xlsx)
        xlTypePDF = 0
        book.ExportAsFixedFormat(xlTypePDF, output_file)
        print(xlsx)
    app.Quit()


# pyinstaller -F -n Vtian_xlsxToPdf main.py
if __name__ == '__main__':
    curr_dir_path = os.getcwd()
    to_pdf(curr_dir_path)
    MegerPDF.merge_all_pdf(curr_dir_path)
