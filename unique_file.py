
import os,time
def main():
    dir = os.getcwd()
    file_names = []
    for file in os.listdir(dir):
        if file.endswith('.jpg'):
            file_names.append(file)

    # 用于存储唯一的字符部分
    unique_characters = set()
    # 用于存储去重后的文件名
    unique_file_names = []

    for file_name in file_names:
        # 在 "ts" 后面截取字符部分
        if "TS" in file_name:
            ts1 = file_name.split("_")[1]
            character = ts1.split('TS')[1]
            # 如果字符部分还没有出现过，则将其添加到 unique_characters 集合中，并将文件名添加到 unique_file_names 列表中
            if character not in unique_characters:
                unique_characters.add(character)
                unique_file_names.append(file_name)
    print("\n---------开始处理文件----------")
    time.sleep(2)
    i=0
    for file in os.listdir(dir):
        if file.endswith('.jpg'):
            if "TS" in file:
                if file not in unique_file_names:
                    os.remove(file)
                    print(f'\n删除文件：{file}')
                    i+=1
                    time.sleep(1)
    print("\n---------处理完成----------")

    print(f"\n---------共删除文件数量：{i}----------")
    print(f"\n---------10秒后窗口将关闭----------")
    time.sleep(10)
# pyinstaller -F -n Vtian_Unique_file --icon=favicon.ico --distpath ./unique_file   unique_file.py
if __name__ == '__main__':
    main()