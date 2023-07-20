import pathlib
import time

import win32com.client
import os
import meger_pdf
import PySimpleGUI as sg
import pc_auth


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


def get_xlsx_files(directory):
    files = []
    for file in os.listdir(directory):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            files.append(file)
    return files


def get_pdf_files(directory):
    files = []
    for file in os.listdir(directory):
        if file.endswith('.pdf') or file.endswith('.PDF'):
            files.append(file)
    return files


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
        time.sleep(2)
    p1.close()


def convertWindow():
    auth_info = pc_auth.readAuthInfo()

    # 创建主窗口布局
    layout1 = [[sg.Text('选择的目录: '), sg.Text('', key='-DIRECTORY1-')],
               [sg.Button('选择目录', key='-CHOOSE_DIR1-')],
               [sg.Text('目录中的 .pdf 文件:')],
               [sg.Listbox(values=[], size=(60, 10), key='-LISTBOX1-', select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
               [sg.Button('确认转换所选文件', key='-check1-')],
               [sg.Output(size=(60, 10), key='-OUTPUT1-')]
               ]

    layout2 = [[sg.Text('选择的目录: '), sg.Text('', key='-DIRECTORY2-')],
               [sg.Button('选择目录', key='-CHOOSE_DIR2-')],
               [sg.Text('目录中的 .xlsx 文件:')],
               [sg.Listbox(values=[], size=(60, 10), key='-LISTBOX2-', select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
               [sg.Button('确认转换所选文件', key='-check1-')],
               [sg.Output(size=(60, 10), key='-OUTPUT2-')]
               ]
    # 创建标签页元素
    tab_group_layout = [
        [sg.Tab('pdf合并', layout1, key='-TAB1-')],
        [sg.Tab('xlsx转pdf且合并', layout2, key='-TAB2-')]
    ]
    # 创建布局
    layout = [
        [sg.TabGroup(tab_group_layout)],
    ]
    tips = ""
    if auth_info["is_permanent"] is False and auth_info["number_of_times"] > -1:
        tips = f': 试用次数剩余{auth_info["number_of_times"]}次'
    # 设定全局图标
    sg.set_global_icon('favicon.ico')
    # 创建主窗口
    window = sg.Window(f'Vtian 转换 {tips}', layout=layout)
    # 当前tab栏
    currTab = 1
    # 事件循环
    while True:

        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, '退出'):
            print(11)
            break

        if auth_info["is_permanent"] is False and event in ('选择目录', "确认转换所选文件"):
            sg.popup('试用次数已用完\n请添加微信：ituserxxx 咨询', title='提示', modal=True)
            continue


        if event == '-CHOOSE_DIR1-':
            currTab = 1
            window.disable()
            folder = sg.popup_get_folder('选择目录', no_window=True)
            window['-DIRECTORY1-'].update(folder)
            window.enable()
            choose_files = get_pdf_files(folder)
            window[f'-LISTBOX{currTab}-'].update(values=choose_files)
            continue
        if event == '-CHOOSE_DIR2-':
            currTab = 2
            window.disable()
            folder = sg.popup_get_folder('选择目录', no_window=True)
            window['-DIRECTORY2-'].update(folder)
            window.enable()
            choose_files = get_xlsx_files(folder)
            window[f'-LISTBOX{currTab}-'].update(values=choose_files)
            continue

        if event in ('-TAB1-', '-TAB2-'):
            currTab = event[4]
            window['-OUTPUT1-'].update("")
            window['-LISTBOX1-'].update("")
            window['-OUTPUT2-'].update("")
            window['-LISTBOX2-'].update("")
            continue
        if event.startswith("-check"):
            selected_files = values[f'-LISTBOX{currTab}-']

            if len(selected_files) == 0:
                sg.popup("没有选择任何文件")
                continue
            root_dir = window[f'-DIRECTORY{currTab}-'].get()
            handle_files = []
            for file in selected_files:
                window[f'-OUTPUT{currTab}-'].update(file + '\n', append=True)
                handle_files.append(os.path.join(root_dir, file))
            if len(handle_files) == 1:
                sg.popup("只有一个文件不需要转换吧~~")
                continue
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            if currTab == 1:
                result = meger_pdf.MergeSomeFileToPDF(desktop_path, handle_files, False)
                window[f'-OUTPUT{currTab}-'].update(result, append=True)
            if currTab == 2:
                result = meger_pdf.MergeSomeFileToPDF(desktop_path, convert_some_file_to_pdf(handle_files), isDelOrigin=True)
                window[f'-OUTPUT{currTab}-'].update(result, append=True)
            if auth_info["is_permanent"] is False:
                auth_info = pc_auth.descNumberOfTimes(1)
                window.set_title(f'Vtian 转换: 试用次数剩余{auth_info["number_of_times"]}次')
    window.close()


# pip freeze > requirements.txt
# pyinstaller -F -n Vtian_xlsxToPdf --icon=favicon.ico --distpath ./unique_file  --noconsole  main.py
if __name__ == '__main__':
    convertWindow()
