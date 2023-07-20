import tkinter as tk
from tkinter import messagebox
import tkinter.filedialog  as filedialog
import os
from tkinter import *

window = tk.Tk()
# 设置窗口title
window.title('Vtian')

# 设置窗口大小变量
width = 600
height = 600
# 窗口居中，获取屏幕尺寸以计算布局参数，使窗口居屏幕中央
screenwidth = window.winfo_screenwidth()
screenheight = window.winfo_screenheight()
size_geo = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
window.geometry(size_geo)


def show1():
    # 若内容是文字则以字符为单位，图像则以像素为单位
    tk.Label(
        window,
        text="网址：c.biancheng.net",
        font=('宋体', 20, 'bold italic'),
        bg="#7CCD7C",
        # 设置标签内容区大小
        width=30, height=5,
        # 设置填充区距离、边框宽度和其样式（凹陷式）
        padx=10, pady=15,
        borderwidth=2,
        relief="sunken"
    ).pack()


# 设置回调函数
def callback():
    print("click me!")


def login(win):
    # 将俩个标签分别布置在第一行、第二行
    tk.Label(win, text="账号：").grid(row=0)
    tk.Label(win, text="密码：").grid(row=1)
    # 创建输入框控件
    e1 = tk.Entry(win)
    # 以 * 的形式显示密码
    e2 = tk.Entry(win, show='*')
    e1.grid(row=0, column=1, padx=10, pady=5)
    e2.grid(row=1, column=1, padx=10, pady=5)

    # 编写一个简单的回调函数
    def login():
        messagebox.showinfo('欢迎您到来')

    # 使用 grid()的函数来布局，并控制按钮的显示位置
    tk.Button(win, text="登录", width=10, command=login).grid(row=3, column=0, sticky="w", padx=10, pady=5)
    tk.Button(win, text="退出", width=10, command=win.quit).grid(row=3, column=1, sticky="e", padx=10, pady=5)


# 使用按钮控件调用函数
# tk.Button(window, text="点击执行回调函数", command=callback).pack()
# show1()
# login(window)
# 定义一个处理文件的相关函数

def fram_left():
    # 创建一个frame窗体对象，用来包裹标签
    frame = tk.Frame(window, relief=SUNKEN, borderwidth=2, width=450, height=250)

    # 设置标签4
    Label4 = tk.Label(frame, text="位置4", bg='gray', fg='white')
    # 设置水平起始位置相对于窗体水平距离的0.01倍，垂直的绝对距离为80，并设置高度为窗体高度比例的0.5倍，宽度为80
    Label4.place(relx=0.01, y=80, relheight=0.4, width=80)


    t1 = tk.Button(frame, text="标签1", command=show_1)
    t1.place(relx=0.04, y=100, relheight=0.08)
    # t2 = tk.Button(frame, text="标签2", width=10, command=show_2)
    # t2.place(relx=0.01, y=40, relheight=0.2, width=5,height=10)
    # 在水平、垂直方向上填充窗体
    frame.pack(side=TOP, fill=BOTH, expand=1)


def show_1():
    print("111")
    # 在主窗口上添加一个frame控件
    frame1 = tk.Frame(window)
    frame1.pack()
    # 在frame_1上添加另外两个frame， 一个在靠左，一个靠右
    # 左侧的frame
    frame_left = tk.Frame(frame1)

    def askfile():
        # 从本地选择一个文件，并返回文件的目录
        path = filedialog.askdirectory()
        if path != '':
            lb.config(text=path)
        else:
            lb.config(text='您没有选择任何目录')

        files = [file for file in os.listdir(path) if file.endswith(".xlsx")]
        print(files)

    btn = tk.Button(frame_left, text='选择目录', relief="raised", command=askfile)
    btn.grid(row=3, column=1, sticky="n")

    lb = tk.Label(frame_left, text='请选择目录', bg='#87CEEB')
    lb.grid(row=3, column=2, padx=10, pady=5)

    frame_left.pack(side=tk.LEFT, anchor="s")


def show_2():
    print("222")


fram_left()

# # 如使用该函数则窗口不能被拉伸
# window.resizable(0,0)
# # 改变背景颜色
# window.config(background="#6fb765")
# # 设置窗口处于顶层
# window.attributes('-topmost',True)
# # 设置窗口的透明度
# window.attributes('-alpha',1)
# # 设置窗口被允许最大调整的范围，与resizble()冲突
# window.maxsize(600,600)
# # 设置窗口被允许最小调整的范围，与resizble()冲突
# window.minsize(600,600)
# 更改左上角窗口的的icon图标,加载C语言中文网logo标
window.iconbitmap('favicon.ico')

# 进入主循环，显示主窗口
window.mainloop()
