import pathlib
import time

import win32com.client
import os
import meger_pdf
import PySimpleGUI as sg

import pc_auth
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
        time.sleep(2)
    p1.close()



def convertWindow():
    auth_info = pc_auth.readAuthInfo()

    # 创建主窗口布局
    layout = [[sg.Text('选择的目录: '), sg.Text('', key='-DIRECTORY-')],
              [sg.Button('选择目录')],
              [sg.Text('目录中的 .xlsx 文件:')],
              [sg.Listbox(values=[], size=(60, 10), key='-LISTBOX-', select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
              [sg.Button('确认转换所选文件')],
              [sg.Output(size=(60, 10), key='-OUTPUT-')]]

    tips =""
    if auth_info["is_permanent"] is False:
        tips  =  f': 试用次数剩余{auth_info["number_of_times"]}次'
    # 创建主窗口
    window = sg.Window(f'Vtian 转换 {tips}', layout)
    # 事件循环
    while True:

        event, values = window.read()
        if auth_info["is_permanent"] is False and auth_info["number_of_times"] == 0 and event in ('选择目录',"确认转换所选文件"):
            sg.popup('试用次数已用完\n请添加微信：ituserxxx 咨询', title='提示',modal=True)
            continue
        if event in (sg.WINDOW_CLOSED, '退出'):
            break


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
            auth_info = pc_auth.descNumberOfTimes(1)
            window.set_title(f'Vtian 转换: 试用次数剩余{auth_info["number_of_times"]}次')

    window.close()



if __name__ == '__main__':
    convertWindow()
