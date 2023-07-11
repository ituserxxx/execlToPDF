import pathlib
import win32com.client
import os
import MegerPDF
import PySimpleGUI as sg


def get_xlsx_files(root_dir):
    xlsx_files = []
    for file in os.listdir(root_dir):
        if file.endswith(".xlsx"):
            xlsx_files.append(os.path.join(root_dir, file))
    return xlsx_files

def to_pdf(curr_dir_path):
    root_dir = curr_dir_path.strip('\"')
    xlsx_files = get_xlsx_files(root_dir)

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


# 创建文件选择对话框
def select_directory():
    layout = [[sg.Text('选择一个目录')],
              [sg.Input(), sg.FolderBrowse()],
              [sg.OK(), sg.Cancel()]]

    window = sg.Window('选择目录', layout)

    while True:
        event, values = window.read()

        if event == 'OK' or event == sg.WINDOW_CLOSED:
            break

    window.close()
    return values[0]


def main():
    directory = select_directory()

    # 创建主窗口布局
    layout = [[sg.Text(f'选择的目录: {directory}')],
              [sg.Button('确认目录')],
              [sg.Text('目录中的 .xlsx 文件:')],
              [sg.Listbox(values=[], size=(60, 10), key='-LISTBOX-', select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
              [sg.Button('转换'), sg.Button('退出')],
              [sg.Output(size=(60, 10), key='-OUTPUT-')]]

    window = sg.Window('目录选择器', layout)

    # 事件循环
    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == '退出':
            break

        if event == '确认目录':
            xlsx_files = get_xlsx_files(directory)
            window['-LISTBOX-'].update(values=xlsx_files)

        if event == '转换':
            selected_files = values['-LISTBOX-']

            # 输出被选中的文件名
            for file in selected_files:
                print(file)

    window.close()

    sg.Window('文件列表', layout).read()
if __name__ == '__main__':
    main()


# pyinstaller -F -n Vtian_xlsxToPdf main.py
# if __name__ == '__main__':
#     curr_dir_path = os.getcwd()
#     to_pdf(curr_dir_path)
#     MegerPDF.merge_all_pdf(curr_dir_path)
