import pathlib
import threading

import win32com.client
import os
import meger_pdf
import PySimpleGUI as sg


def get_xlsx(root_dir):
    xlsx_files = []
    for file in os.listdir(root_dir):
        if file.endswith(".xlsx"):
            xlsx_files.append(os.path.join(root_dir, file))
    return xlsx_files


def convert_some_file_to_pdf(xlsx_files):
    # root_dir = curr_dir_path.strip('\"')
    # xlsx_files = get_xlsx(root_dir)
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    new_pdf_list = []
    for i in xlsx_files:
        xlsx = pathlib.Path(i)
        xlsx_dir = xlsx.parent
        xlsx_dir = str(xlsx_dir)
        basename = xlsx.stem
        basename = str(basename)
        output_file = xlsx_dir + "/" + basename + ".pdf"
        book = app.Workbooks.Open(xlsx)
        xlTypePDF = 0  # 固定只转换 xlsx 的第一页签：sheet1
        book.ExportAsFixedFormat(xlTypePDF, output_file)
        new_pdf_list.append(output_file)
    app.Quit()
    return new_pdf_list


# 创建文件选择对话框
def select_directory():
    layout = [[sg.Text('选择一个目录')],
              [sg.Input(), sg.FolderBrowse(key='-FOLDER-')],
              [sg.OK(), sg.Cancel()]]

    window = sg.Window('选择目录', layout)

    while True:
        event, values = window.read()
        if event in ('OK', sg.WINDOW_CLOSED):
            break

    window.close()
    return values['-FOLDER-']


# 获取目录中以 .xlsx 结尾的文件名称
def get_xlsx_files(directory):
    xlsx_files = []
    for file in os.listdir(directory):
        if file.endswith('.xlsx'):
            xlsx_files.append(file)
    return xlsx_files


def sub_window_thread():
    # 创建子窗口并运行事件循环
    lay = [[sg.Text('这是一个自定义大小的弹窗')],
           [sg.Multiline('', size=(30, 10))],
           [sg.Button('确定')],
           [sg.Button('取消')]
           ]
    # 创建弹窗并运行事件循环
    p1 = sg.Window('自定义大小的弹窗示例', lay)
    event, values = p1.read()
    while True:
        if event in (sg.WINDOW_CLOSED, '退出'):
            break
        if event == "取消":
            break
    p1.close()

def convertWindow():
    # 创建主窗口布局
    layout = [[sg.Text('选择的目录: '), sg.Text('', key='-DIRECTORY-')],
              [sg.Button('选择目录')],
              [sg.Text('目录中的 .xlsx 文件:')],
              [sg.Listbox(values=[], size=(60, 10), key='-LISTBOX-', select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
              [sg.Button('确认转换所选文件')],
              [sg.Output(size=(60, 10), key='-OUTPUT-')]]

    # 创建主窗口
    # window = sg.Window('Vtian 转换 : 试用次数剩余3次', layout)
    window = sg.Window('Vtian 转换 ', layout)

    use_time = 0
    # 事件循环
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, '退出'):
            break
        if use_time == 0 and event in ('选择目录',"确认转换所选文件"):
            window.disable()
            t = threading.Thread(target=sub_window_thread)
            t.start()
            # sg.popup('试用次数已用完', title='提示', custom_text=("立即添加咨询", "取消"))

            window.enable()
            continue

        if event == '选择目录':
            directory = select_directory()
            window['-DIRECTORY-'].update(directory)

            choose_files = get_xlsx_files(directory)
            window['-LISTBOX-'].update(values=choose_files)

            window['-OUTPUT-'].update(value="")

        if event == '确认转换所选文件':
            window['-OUTPUT-'].update(value="")
            selected_files = values['-LISTBOX-']
            if len(selected_files) == 0:
                sg.popup("没有选择任何文件")
                continue
            root_dir = window['-DIRECTORY-'].get()
            xlsx_files = []
            for file in selected_files:
                print(file)
                xlsx_files.append(os.path.join(root_dir, file))

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            meger_pdf.MergeSomeFileToPDF(desktop_path, convert_some_file_to_pdf(xlsx_files))

    window.close()


def main():
    # 创建第一个页面的布局
    layout1 = [[sg.Text('这是第一个页面')], [sg.Button('切换到第二个页面')]]

    # 创建第二个页面的布局
    layout2 = [[sg.Text('这是第二个页面')], [sg.Button('切换到第一个页面')]]

    # 创建一个窗口并将第一个页面设置为默认显示
    window = sg.Window('页面切换示例', layout1)
    sg.popup("当前试用次数剩余3")
    while True:
        event, _ = window.read()

        # 根据事件类型进行相应的操作
        if event == sg.WINDOW_CLOSED:
            break
        elif event == '切换到第二个页面':
            window.close()
            window = sg.Window('页面切换示例', layout2)
        elif event == '切换到第一个页面':
            window.close()
            window = sg.Window('页面切换示例', layout1)

    window.close()


if __name__ == '__main__':
    main()
