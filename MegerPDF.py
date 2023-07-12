# -*- coding:utf-8*-
# 利用PyPDF2模块合并同一文件夹下的所有PDF文件
# 只需修改存放PDF文件的文件夹变量：file_dir 和 输出文件名变量: outfile

import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import datetime


# 使用os模块的walk函数，搜索出指定目录下的全部PDF文件
# 获取同一目录下的所有PDF文件的绝对路径
def getFileName(filedir):
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []

# 合并一些文件到一个PDF文件
def MergeSomeFileToPDF(root_dir,some_pdf_fileName):
    outfile = datetime.datetime.now().strftime(f'{root_dir}\\newPDF_Vtian_%Y-%m-%d_%H_%M_%S.pdf')  # 输出的PDF文件的名称
    output = PdfFileWriter()
    outputPages = 0
    rm_file_list =[]
    for pdf_file in some_pdf_fileName:
        # 读取源PDF文件
        f = open(pdf_file, "rb")
        input = PdfFileReader(f)
        # 获得源PDF文件中页面总数
        pageCount = input.getNumPages()
        outputPages += 1

        # 分别将page添加到输出output中
        for iPage in range(pageCount):
            output.addPage(input.getPage(iPage))
            rm_file_list.append([f,pdf_file])
            break


    # 写入到目标PDF文件
    outputStream = open(outfile, "wb")
    output.write(outputStream)
    outputStream.close()

    for ff in rm_file_list:
        ff[0].close()
        os.remove(ff[1])

    print(f"PDF文件合并完成！！！\n总页数:{outputPages}")
    print(f"合并后的文件存放位置：\n {outfile}")


# 合并同一目录下的所有PDF文件
def MergePDFto(filepath, outfile):
    output = PdfFileWriter()
    outputPages = 0
    pdf_fileName = getFileName(filepath)
    rm_file_list =[]
    if pdf_fileName:
        for pdf_file in pdf_fileName:
            # 读取源PDF文件
            f = open(pdf_file, "rb")
            input = PdfFileReader(f)
            # 获得源PDF文件中页面总数
            pageCount = input.getNumPages()
            outputPages += 1

            # 分别将page添加到输出output中
            for iPage in range(pageCount):
                output.addPage(input.getPage(iPage))
                rm_file_list.append([f,pdf_file])
                break

        print("合并后的总页数:%d." % outputPages)
        # 写入到目标PDF文件
        outputStream = open(outfile, "wb")
        output.write(outputStream)
        outputStream.close()

        for ff in rm_file_list:
            ff[0].close()
            os.remove(ff[1])

        print("PDF文件合并完成！")
    else:
        print("没有可以合并的PDF文件！")

# 主函数
def merge_all_pdf(file_dir):
    # file_dir = r'E:\Cheats' # 存放PDF的原文件夹
    outfile = datetime.datetime.now().strftime('newPDF_%Y-%m-%d_%H_%M_%S.pdf')  # 输出的PDF文件的名称
    MergePDFto(file_dir, outfile)



