import pathlib
import win32com.client
import os
import MegerPDF

def all_files(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def to_pdf(asb_cwd ):
    root_dir = asb_cwd.strip('\"')
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    for i in all_files(root_dir):
        xlsx = pathlib.Path(i)
        if xlsx.suffix == ".xlsx":
            print(i)
            xlsx_dir = xlsx.parent
            xlsx_dir = str(xlsx_dir)
            basename = xlsx.stem
            basename = str(basename)
            output_file = xlsx_dir + "/" + basename + ".pdf"
            book = app.Workbooks.Open(xlsx)
            xlTypePDF = 0
            book.ExportAsFixedFormat(xlTypePDF, output_file)
    app.Quit()

# pyinstaller -F -n Vtian_xlsxToPdf main.py
if __name__ == '__main__':
    asb_cwd = os.getcwd()
    to_pdf(asb_cwd)
    MegerPDF.merge_all_pdf(asb_cwd)



